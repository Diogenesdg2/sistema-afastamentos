import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from tkcalendar import DateEntry
from datetime import datetime, timedelta
from perda_ferias import emitir_termo_ferias
from recurso import TelaRecurso
import json
import os
import csv
import sqlite3
import base64
import platform
import subprocess
from database import (
    conectar_db, calcular_fim_aquisitivo, calcular_dias_beneficio_aquisitivo,
    sincronizar_aquisitivo_funcionario, obter_ids_mesmo_funcionario
)

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib import colors

CHAVE_MESTRE = "VimanSeguranca"

def encriptar_senha(pwd):
    if not pwd: return ""
    if pwd.startswith("ENC::"): return pwd
    xored = ''.join(chr(ord(c) ^ ord(k)) for c, k in zip(pwd, CHAVE_MESTRE * (len(pwd)//len(CHAVE_MESTRE) + 1)))
    return "ENC::" + base64.b64encode(xored.encode('utf-8')).decode('utf-8')

def desencriptar_senha(encoded):
    if not encoded: return ""
    if not encoded.startswith("ENC::"): return encoded 
    try:
        encoded_str = encoded[5:]
        xored = base64.b64decode(encoded_str.encode('utf-8')).decode('utf-8')
        return ''.join(chr(ord(c) ^ ord(k)) for c, k in zip(xored, CHAVE_MESTRE * (len(xored)//len(CHAVE_MESTRE) + 1)))
    except:
        return encoded

def buscar_cid_base(codigo):
    cod_busca = codigo.strip().upper().replace(".", "")
    if not cod_busca: return ""
    for nome_arquivo in ["cid10.csv", "CID10.CSV", "cid-10.csv", "CID-10.CSV", "cid10_local.json"]:
        if not os.path.exists(nome_arquivo): continue
        if nome_arquivo.endswith(".json"):
            try:
                with open(nome_arquivo, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                    for item in dados:
                        if str(item.get("codigo", item.get("cid", ""))).strip().upper().replace(".", "") == cod_busca:
                            return item.get("nome", item.get("descricao", ""))
            except Exception: pass
            continue
        for enc in ["latin-1", "utf-8-sig", "utf-8", "cp1252"]:
            try:
                with open(nome_arquivo, "r", encoding=enc) as f:
                    primeira_linha = f.readline()
                    delim = ";" if ";" in primeira_linha else ("," if "," in primeira_linha else "\t")
                    f.seek(0)
                    leitor = csv.reader(f, delimiter=delim)
                    header = next(leitor, None)
                    desc_idx = -1
                    if header:
                        header_norm = [col.strip().upper() for col in header]
                        if "DESCRICAO" in header_norm: desc_idx = header_norm.index("DESCRICAO")
                        elif "NOME" in header_norm: desc_idx = header_norm.index("NOME")
                    for linha in leitor:
                        if not linha or not linha[0]: continue
                        if linha[0].strip().upper().replace(".", "") == cod_busca:
                            if desc_idx != -1 and len(linha) > desc_idx and linha[desc_idx].strip(): return linha[desc_idx].strip()
                            if len(linha) > 4 and linha[4].strip(): return linha[4].strip()
                            if len(linha) > 1 and linha[1].strip(): return linha[1].strip()
                            if len(linha) > 5 and linha[5].strip(): return linha[5].strip()
                    break
            except UnicodeDecodeError: continue
            except Exception: continue
    return ""

class DateEntryLivre(DateEntry):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.unbind("<<FocusOut>>")
        self.unbind("<FocusOut>")

    def _validate_date(self):
        if not self.get(): return True
        return super()._validate_date()

    def parse_date(self, text):
        if not text: return datetime.now().date()
        try: return super().parse_date(text)
        except Exception: return datetime.now().date()

    def drop_down(self):
        if not self.get():
            self.set_date(datetime.now())
            super().drop_down()
            self.delete(0, tk.END)
        else:
            super().drop_down()

class TelaManutencao:
    def __init__(self, parent, id_registro, callback_atualizar):
        self.id_registro = id_registro
        self.callback_atualizar = callback_atualizar
        self.dias_beneficio_aquisitivo = 0
        
        self.janela = tk.Toplevel(parent)
        self.janela.withdraw() 
        
        self.dados = self.buscar_dados()
        if not self.dados:
            messagebox.showerror("Erro", "Registo não encontrado.", parent=self.janela)
            self.janela.destroy()
            return

        nome_func = self.dados['nome_empregado'] or "Sem Nome"
        self.janela.title(f"Ficha Completa do Afastamento | ID: {self.id_registro} - {nome_func}")
        self.janela.geometry("1200x840") 
        self.janela.grab_set()
        self.janela.resizable(False, False)
        self.janela.configure(bg="#F3F4F6")

        self.map_analises_mae = {} 

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook", background="#F3F4F6", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[15, 8], background="#E5E7EB", foreground="#374151")
        style.map("TNotebook.Tab", background=[("selected", "#FFFFFF")], foreground=[("selected", "#1D4ED8")])
        style.configure("TFrame", background="#FFFFFF")
        style.configure("TLabelframe", background="#FFFFFF", borderwidth=2)
        style.configure("TLabelframe.Label", background="#FFFFFF", font=("Segoe UI", 10, "bold"), foreground="#2563EB")

        self.notebook = ttk.Notebook(self.janela)
        self.notebook.pack(fill="both", expand=True, padx=20, pady=(20, 10))

        self.aba_resumo = ttk.Frame(self.notebook)
        self.aba_inss = ttk.Frame(self.notebook)
        self.aba_concessao = ttk.Frame(self.notebook)
        self.aba_retorno = ttk.Frame(self.notebook)

        self.notebook.add(self.aba_resumo, text=" 👤 Dados Básicos ")
        self.notebook.add(self.aba_inss, text=" ⚖️ Acompanhamento ")
        self.notebook.add(self.aba_concessao, text=" 🏛️ INSS ")
        self.notebook.add(self.aba_retorno, text=" 🔄 Retorno ")

        self.construir_aba_resumo()
        self.construir_aba_acompanhamento()
        self.construir_aba_inss()
        self.construir_aba_retorno()

        frame_botoes = tk.Frame(self.janela, bg="#F3F4F6")
        frame_botoes.pack(fill="x", padx=20, pady=(0, 20))

        btn_salvar = tk.Button(frame_botoes, text="💾 Salvar Todas as Alterações", command=self.salvar_alteracoes, 
                               bg="#10B981", fg="white", font=("Segoe UI", 11, "bold"), padx=40, pady=8, relief="flat", cursor="hand2")
        btn_salvar.pack(side="left", padx=(0, 10))
        
        # O botão CADASTRAR RECURSO fica ao lado do Salvar
        btn_recurso = tk.Button(frame_botoes, text="📋 CADASTRAR Recurso/Revisão/Contestação", command=self.abrir_modal_recurso, 
                               bg="#F59E0B", fg="white", font=("Segoe UI", 11, "bold"), padx=20, pady=8, relief="flat", cursor="hand2")
        btn_recurso.pack(side="left")

        self.preencher_dados()
        self.janela.deiconify() 

    def criar_calendario(self, parent_frame, width=12, cor_fundo='#3B82F6'):
        return DateEntryLivre(parent_frame, width=width, background=cor_fundo, foreground='white', borderwidth=0, 
                             date_pattern='dd/mm/yyyy', state="normal", font=("Segoe UI", 10))

    def atualizar_termometro_ferias(self, event=None):
        dt_de = self.ent_aq_de.get().strip()
        dt_ate = calcular_fim_aquisitivo(dt_de)
        self.ent_aq_ate.config(state="normal")
        self.ent_aq_ate.delete(0, tk.END)
        if dt_ate:
            self.ent_aq_ate.insert(0, dt_ate)
        self.ent_aq_ate.config(state="readonly")

        if self.var_ctrl_ferias.get() == 0 or not dt_de or not dt_ate:
            self.dias_beneficio_aquisitivo = 0
            self.lbl_termometro_dias.config(text="Controle Desativado", fg="#6B7280", bg="#F3F4F6")
            self.lbl_termometro_status.config(text="Marque 'Controlar Período Aquisitivo' para monitorizar os 180 dias.", fg="#6B7280")
            if hasattr(self, 'lbl_inss_ferias'):
                self.lbl_inss_ferias.config(text="🏖️ Controle de Férias Desativado para este registo", fg="#6B7280", bg="#F3F4F6")
                self.f_inss_ferias.config(bg="#F3F4F6")
            return

        dias, qtd_reg = calcular_dias_beneficio_aquisitivo(
            self.ent_cpf.get(), self.ent_cod_empresa.get(), self.ent_cod_func.get(),
            dt_de, dt_ate, id_atual=self.id_registro
        )
        self.dias_beneficio_aquisitivo = dias

        if dias >= 180:
            cor_fg, cor_bg = "#DC2626", "#FEE2E2"
            msg_st = f"🚨 PERDA DO DIREITO A FÉRIAS! ({dias} dias concedidos em {qtd_reg} afastamento(s))"
        elif dias >= 150:
            cor_fg, cor_bg = "#D97706", "#FEF3C7"
            msg_st = f"⚠️ Atenção: Faltam {180 - dias} dias de benefício para perder as férias ({qtd_reg} reg.)"
        else:
            cor_fg, cor_bg = "#15803D", "#DCFCE7"
            msg_st = f"🟢 Direito a férias mantido (Restam {180 - dias} dias até ao limite | {qtd_reg} reg.)"

        self.lbl_termometro_dias.config(text=f"{dias} / 180 dias concedidos", fg=cor_fg, bg=cor_bg)
        self.lbl_termometro_status.config(text=msg_st, fg=cor_fg)

        if hasattr(self, 'lbl_inss_ferias'):
            self.lbl_inss_ferias.config(text=f"🏖️ Resumo do Período Aquisitivo ({dt_de} a {dt_ate}): {dias} / 180 dias concedidos pelo INSS", fg=cor_fg, bg=cor_bg)
            self.f_inss_ferias.config(bg=cor_bg)

        self.verificar_virada_aquisitivo_retorno()

    def abrir_detalhes_ferias(self, event=None):
        if self.var_ctrl_ferias.get() == 0:
            return messagebox.showinfo("Aviso", "O controle de férias está desativado para este registo.", parent=self.janela)

        aq_de_str = self.ent_aq_de.get().strip()
        aq_ate_str = self.ent_aq_ate.get().strip()
        if not aq_de_str or not aq_ate_str:
            return messagebox.showwarning("Aviso", "Defina o Início do Período Aquisitivo primeiro.", parent=self.janela)

        conn = conectar_db()
        cursor = conn.cursor()
        
        ids = obter_ids_mesmo_funcionario(
            cursor, self.ent_cpf.get().strip(), self.ent_cod_empresa.get().strip(), 
            self.ent_cod_func.get().strip(), self.id_registro
        )
        
        if not ids:
            conn.close()
            return

        qmarks = ",".join("?" for _ in ids)
        periodos = []
        
        cursor.execute(f"SELECT id, id_afastamento, concessao_de, concessao_ate FROM analises_inss WHERE id_afastamento IN ({qmarks}) AND situacao IN ('Concedido', 'Concedida')", ids)
        for r in cursor.fetchall():
            if r[2] and r[3]: periodos.append({'tipo': 'Análise Doc.', 'id_item': r[0], 'id_af': r[1], 'de': r[2], 'ate': r[3]})
            
        cursor.execute(f"SELECT id, id_afastamento, concessao_de, concessao_ate FROM pericias_inss WHERE id_afastamento IN ({qmarks}) AND situacao IN ('Concedido', 'Concedida')", ids)
        for r in cursor.fetchall():
            if r[2] and r[3]: periodos.append({'tipo': 'Perícia Pres.', 'id_item': r[0], 'id_af': r[1], 'de': r[2], 'ate': r[3]})
            
        conn.close()

        aq_de_dt = datetime.strptime(aq_de_str, "%d/%m/%Y").date()
        aq_ate_dt = datetime.strptime(aq_ate_str, "%d/%m/%Y").date()

        detalhes = []
        for p in periodos:
            try:
                p_de = datetime.strptime(p['de'], "%d/%m/%Y").date()
                p_ate = datetime.strptime(p['ate'], "%d/%m/%Y").date()
                
                ini_efetivo = max(p_de, aq_de_dt)
                fim_efetivo = min(p_ate, aq_ate_dt)
                
                if ini_efetivo <= fim_efetivo:
                    dias = (fim_efetivo - ini_efetivo).days + 1
                    detalhes.append({
                        'desc': f"{p['tipo']} #{p['id_item']} (Af. ID {p['id_af']})",
                        'original': f"{p['de']} a {p['ate']}",
                        'contado': f"{ini_efetivo.strftime('%d/%m/%Y')} a {fim_efetivo.strftime('%d/%m/%Y')}",
                        'dias': dias
                    })
            except ValueError:
                pass
                
        jan = tk.Toplevel(self.janela)
        jan.title("Extrato de Dias Concedidos (Férias)")
        jan.geometry("780x420")
        jan.grab_set()
        jan.configure(bg="#F3F4F6")
        
        ttk.Label(jan, text=f"Extrato do Período Aquisitivo: {aq_de_str} a {aq_ate_str}", font=("Segoe UI", 12, "bold"), background="#F3F4F6", foreground="#0284C7").pack(pady=(15, 5))
        
        tree = ttk.Treeview(jan, columns=("Origem", "Período Original", "Período Contado", "Dias"), show="headings", height=10)
        tree.heading("Origem", text="Origem do Benefício")
        tree.heading("Período Original", text="Período Total")
        tree.heading("Período Contado", text="Validado no Aquisitivo")
        tree.heading("Dias", text="Dias")
        
        tree.column("Origem", width=220, anchor="w")
        tree.column("Período Original", width=150, anchor="center")
        tree.column("Período Contado", width=150, anchor="center")
        tree.column("Dias", width=80, anchor="center")
        tree.pack(fill="x", padx=20, pady=5)
        
        for d in detalhes:
            tree.insert("", "end", values=(d['desc'], d['original'], d['contado'], f"{d['dias']} dias"))
            
        tk.Label(jan, text=f"Total Deduplicado (Dias Únicos Restantes): {self.dias_beneficio_aquisitivo} dias", font=("Segoe UI", 12, "bold"), bg="#F3F4F6", fg="#DC2626").pack(pady=10)
        tk.Button(jan, text="Fechar Extrato", command=jan.destroy, bg="#3B82F6", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=6).pack(pady=(0, 10))

    def verificar_virada_aquisitivo_retorno(self, event=None):
        if not hasattr(self, 'ent_ret_data') or not hasattr(self, 'ent_novo_aq_de'):
            return

        ret_str = self.ent_ret_data.get().strip()
        motivo = self.combo_motivo.get()
        is_aposentado = str(getattr(self, 'var_aposentado', tk.IntVar()).get()) in ("1", "Sim", "True")
        
        try:
            total_dias = int(self.ent_dias.get().strip())
        except (ValueError, TypeError):
            total_dias = 0

        self.btn_termo_ferias.pack_forget()

        if motivo in ["Cárcere", "Militar"]:
            if len(ret_str) == 10:
                self.lbl_aviso_perda_retorno.config(
                    text=f"⚖️ Contrato suspenso por {motivo}. O período aquisitivo será complementado pelos dias de afastamento.",
                    fg="#D97706"
                )
                
                try:
                    ret_dt = datetime.strptime(ret_str, "%d/%m/%Y").date()
                    fim_afast_dt = ret_dt - timedelta(days=1)
                    per_de_str = self.ent_per_de.get().strip()
                    
                    if not per_de_str and self.ent_ult_dia.get().strip():
                        ult_dt = datetime.strptime(self.ent_ult_dia.get().strip(), "%d/%m/%Y").date()
                        per_de_str = (ult_dt + timedelta(days=1)).strftime("%d/%m/%Y")
                        self.ent_per_de.config(state="normal")
                        self.ent_per_de.delete(0, tk.END)
                        self.ent_per_de.insert(0, per_de_str)
                        self.ent_per_de.config(state="readonly")

                    if per_de_str:
                        ini_dt = datetime.strptime(per_de_str, "%d/%m/%Y").date()
                        if fim_afast_dt >= ini_dt:
                            total_dias_calc = (fim_afast_dt - ini_dt).days + 1
                            fim_str = fim_afast_dt.strftime("%d/%m/%Y")
                            
                            for widget, val in [(self.ent_per_ate, fim_str), (self.ent_dias, str(total_dias_calc)), (self.ent_soma_atestados, str(total_dias_calc))]:
                                widget.config(state="normal")
                                widget.delete(0, tk.END)
                                widget.insert(0, val)
                                widget.config(state="readonly")

                            aq_de_str = self.ent_aq_de.get().strip()
                            if len(aq_de_str) == 10:
                                from perda_ferias import calcular_meses_dias_ciclo, somar_meses
                                aq_de_dt = datetime.strptime(aq_de_str, "%d/%m/%Y").date()
                                ult_trab_dt = ini_dt - timedelta(days=1)
                                m_antes, d_antes = calcular_meses_dias_ciclo(aq_de_dt, ult_trab_dt)
                                if d_antes >= 30:
                                    m_antes += d_antes // 30
                                    d_antes = d_antes % 30
                                if d_antes == 0:
                                    m_falta, d_falta = max(0, 12 - m_antes), 0
                                else:
                                    m_falta, d_falta = max(0, 11 - m_antes), 30 - d_antes
                                
                                base_ciclo_dt = ret_dt - timedelta(days=1)
                                fim_meses_dt = somar_meses(base_ciclo_dt, m_falta)
                                fim_compl_dt = fim_meses_dt + timedelta(days=d_falta)
                                novo_aq_dt = fim_compl_dt + timedelta(days=1)
                                
                                if not self.ent_novo_aq_de.get().strip():
                                    self.ent_novo_aq_de.delete(0, tk.END)
                                    self.ent_novo_aq_de.insert(0, novo_aq_dt.strftime("%d/%m/%Y"))
                                    self.atualizar_fim_novo_aquisitivo()
                except ValueError:
                    pass

                self.btn_termo_ferias.config(text=f"📄 Emitir Termo ({motivo})")
                self.btn_termo_ferias.pack(pady=(10, 0), anchor="w", padx=10)
            else:
                self.lbl_aviso_perda_retorno.config(
                    text=f"ℹ️ Afastamento por {motivo}. Preencha o Retorno para emitir o termo e calcular a complementação das férias.",
                    fg="#6B7280"
                )

        elif is_aposentado and motivo in ["Doença", "Acidente", "Outros"]:
            if total_dias > 30:
                self.lbl_aviso_perda_retorno.config(
                    text=f"ℹ️ Colaborador Aposentado com afastamento superior a 30 dias ({total_dias} dias). Direito a férias mantido com complementação após o retorno.",
                    fg="#15803D"
                )
                if len(ret_str) == 10:
                    if not self.ent_novo_aq_de.get().strip():
                        try:
                            from perda_ferias import calcular_meses_dias_ciclo, somar_meses
                            aq_de_str = self.ent_aq_de.get().strip()
                            per_de_str = self.ent_per_de.get().strip()
                            if len(aq_de_str) == 10 and len(per_de_str) == 10:
                                aq_de_dt = datetime.strptime(aq_de_str, "%d/%m/%Y").date()
                                ini_dt = datetime.strptime(per_de_str, "%d/%m/%Y").date()
                                ret_dt = datetime.strptime(ret_str, "%d/%m/%Y").date()
                                ult_trab_dt = ini_dt - timedelta(days=1)
                                
                                m_antes, d_antes = calcular_meses_dias_ciclo(aq_de_dt, ult_trab_dt)
                                if d_antes >= 30:
                                    m_antes += d_antes // 30
                                    d_antes = d_antes % 30
                                if d_antes == 0:
                                    m_falta, d_falta = max(0, 12 - m_antes), 0
                                else:
                                    m_falta, d_falta = max(0, 11 - m_antes), 30 - d_antes
                                
                                base_ciclo_dt = ret_dt - timedelta(days=1)
                                fim_meses_dt = somar_meses(base_ciclo_dt, m_falta)
                                fim_compl_dt = fim_meses_dt + timedelta(days=d_falta)
                                novo_aq_dt = fim_compl_dt + timedelta(days=1)
                                
                                self.ent_novo_aq_de.delete(0, tk.END)
                                self.ent_novo_aq_de.insert(0, novo_aq_dt.strftime("%d/%m/%Y"))
                                self.atualizar_fim_novo_aquisitivo()
                        except ValueError:
                            pass

                    self.btn_termo_ferias.config(text="📄 Emitir Termo (Aposentado)")
                    self.btn_termo_ferias.pack(pady=(10, 0), anchor="w", padx=10)
            else:
                self.lbl_aviso_perda_retorno.config(
                    text=f"ℹ️ Colaborador Aposentado com afastamento de {total_dias} dias (inferior ou igual a 30 dias). Sem impacto nas férias.",
                    fg="#6B7280"
                )

        else:
            if self.var_ctrl_ferias.get() == 1 and self.dias_beneficio_aquisitivo >= 180:
                self.lbl_aviso_perda_retorno.config(
                    text=f"🚨 Colaborador acumulou {self.dias_beneficio_aquisitivo} dias concedidos (≥ 180 dias) no período aquisitivo atual.\n"
                         f"Pelo Art. 133, § 2º da CLT, inicia-se um novo período aquisitivo na Data Efetiva de Retorno!",
                    fg="#DC2626"
                )
                if len(ret_str) == 10:
                    self.ent_novo_aq_de.delete(0, tk.END)
                    self.ent_novo_aq_de.insert(0, ret_str)
                    self.atualizar_fim_novo_aquisitivo()
                    
                    self.btn_termo_ferias.config(text="📄 Emitir Termo de Perda de Férias")
                    self.btn_termo_ferias.pack(pady=(10, 0), anchor="w", padx=10)
            else:
                self.lbl_aviso_perda_retorno.config(
                    text="ℹ️ O colaborador possui menos de 180 dias de benefício concedido no ciclo atual (Direito a férias mantido).",
                    fg="#15803D"
                )

    def atualizar_fim_novo_aquisitivo(self, event=None):
        dt_de = self.ent_novo_aq_de.get().strip()
        dt_ate = calcular_fim_aquisitivo(dt_de)
        self.ent_novo_aq_ate.config(state="normal")
        self.ent_novo_aq_ate.delete(0, tk.END)
        if dt_ate:
            self.ent_novo_aq_ate.insert(0, dt_ate)
        self.ent_novo_aq_ate.config(state="readonly")

    def sincronizar_totais_afastamento(self):
        conn = conectar_db()
        cursor = conn.cursor()
        cursor.execute("SELECT data_de, data_ate FROM atestados_historico WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
        atestados = cursor.fetchall()
        
        motivo_atual = self.combo_motivo.get()

        if atestados:
            datas_de, datas_ate = [], []
            for a in atestados:
                try: datas_de.append(datetime.strptime(a[0], "%d/%m/%Y"))
                except: pass
                try: datas_ate.append(datetime.strptime(a[1], "%d/%m/%Y"))
                except: pass
            
            if datas_de and datas_ate:
                min_de, max_ate = min(datas_de), max(datas_ate)
                total_dias = (max_ate - min_de).days + 1
                str_de, str_ate = min_de.strftime("%d/%m/%Y"), max_ate.strftime("%d/%m/%Y")
                
                cursor.execute("UPDATE afastamentos SET periodo_de=?, periodo_ate=?, qtd_dias=? WHERE id=?", (str_de, str_ate, total_dias, self.id_registro))
                
                for widget, val in [(self.ent_per_de, str_de), (self.ent_per_ate, str_ate), (self.ent_dias, str(total_dias))]:
                    widget.config(state="normal")
                    widget.delete(0, tk.END)
                    widget.insert(0, val)
                    widget.config(state="readonly")
        else:
            if motivo_atual in ["Cárcere", "Militar"]:
                ult_trab = self.ent_ult_dia.get().strip()
                if ult_trab:
                    try:
                        dt_inicio_afast = datetime.strptime(ult_trab, "%d/%m/%Y") + timedelta(days=1)
                        str_de = dt_inicio_afast.strftime("%d/%m/%Y")
                        cursor.execute("UPDATE afastamentos SET periodo_de=?, periodo_ate='', qtd_dias=0 WHERE id=?", (str_de, self.id_registro))
                        
                        self.ent_per_de.config(state="normal")
                        self.ent_per_de.delete(0, tk.END)
                        self.ent_per_de.insert(0, str_de)
                        self.ent_per_de.config(state="readonly")
                    except: pass
            else:
                cursor.execute("UPDATE afastamentos SET periodo_de='', periodo_ate='', qtd_dias='' WHERE id=?", (self.id_registro,))
                for obj in [self.ent_per_de, self.ent_per_ate, self.ent_dias]:
                    obj.config(state="normal")
                    obj.delete(0, tk.END)
                    obj.config(state="readonly")

        cursor.execute("SELECT sum(qtd_dias) FROM atestados_historico WHERE id_afastamento=?", (self.id_registro,))
        soma = cursor.fetchone()[0]
        self.ent_soma_atestados.config(state="normal")
        self.ent_soma_atestados.delete(0, tk.END)
        self.ent_soma_atestados.insert(0, str(soma) if soma else "0")
        self.ent_soma_atestados.config(state="readonly")
                
        conn.commit()
        conn.close()
        self.callback_atualizar()

    def construir_aba_resumo(self):
        f = self.aba_resumo
        f.columnconfigure(0, weight=6)
        f.columnconfigure(1, weight=4)
        
        frame_esq = tk.Frame(f, bg="#FFFFFF")
        frame_esq.grid(row=0, column=0, sticky="nsew", padx=20, pady=15)
        frame_dir = tk.Frame(f, bg="#FFFFFF")
        frame_dir.grid(row=0, column=1, sticky="nsew", padx=(0, 20), pady=15)

        fonte_label = ("Segoe UI", 10, "bold")
        fonte_entry = ("Segoe UI", 10)

        tk.Label(frame_esq, text="Empresa:", bg="white", font=fonte_label).grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.ent_empresa = tk.Entry(frame_esq, width=28, font=fonte_entry, relief="solid", bd=1)
        self.ent_empresa.grid(row=0, column=1, sticky="w", pady=(0, 8), padx=(5, 15))

        tk.Label(frame_esq, text="Cód. Empresa:", bg="white", font=fonte_label).grid(row=0, column=2, sticky="e", pady=(0, 8))
        self.ent_cod_empresa = tk.Entry(frame_esq, width=10, font=fonte_entry, relief="solid", bd=1)
        self.ent_cod_empresa.grid(row=0, column=3, sticky="w", pady=(0, 8), padx=5)

        tk.Label(frame_esq, text="Funcionário:", bg="white", font=fonte_label).grid(row=1, column=0, sticky="w", pady=(0, 8))
        self.ent_nome = tk.Entry(frame_esq, width=28, font=fonte_entry, relief="solid", bd=1)
        self.ent_nome.grid(row=1, column=1, sticky="w", pady=(0, 8), padx=(5, 15))

        tk.Label(frame_esq, text="Cód. Func.:", bg="white", font=fonte_label).grid(row=1, column=2, sticky="e", pady=(0, 8))
        self.ent_cod_func = tk.Entry(frame_esq, width=10, font=fonte_entry, relief="solid", bd=1)
        self.ent_cod_func.grid(row=1, column=3, sticky="w", pady=(0, 8), padx=5)

        tk.Label(frame_esq, text="Motivo:", bg="white", font=fonte_label).grid(row=2, column=0, sticky="w", pady=(0, 8))
        self.combo_motivo = ttk.Combobox(frame_esq, values=["Doença", "Acidente", "Cárcere", "Militar", "Maternidade", "Outros"], width=26, state="readonly", font=fonte_entry)
        self.combo_motivo.grid(row=2, column=1, sticky="w", pady=(0, 8), padx=(5, 15))

        tk.Label(frame_esq, text="CIDs Relacionados:", bg="white", font=fonte_label).grid(row=2, column=2, sticky="e", pady=(0, 8))
        f_cid = tk.Frame(frame_esq, bg="white")
        f_cid.grid(row=2, column=3, sticky="w", pady=(0, 8), padx=5)
        self.ent_cid_show = tk.Entry(f_cid, width=12, font=fonte_entry, relief="solid", bd=1, state="readonly", bg="#E5E7EB")
        self.ent_cid_show.pack(side="left")
        tk.Button(f_cid, text="➕ Gerir", command=self.abrir_historico_cids, bg="#3B82F6", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=(5,0))

        f_chk = tk.Frame(frame_esq, bg="white")
        f_chk.grid(row=3, column=0, columnspan=4, sticky="w", pady=(0, 8))
        self.var_aposentado = tk.IntVar()
        tk.Checkbutton(f_chk, text="Aposentado(a) (Inativa aba INSS)", variable=self.var_aposentado, command=self.toggle_aba_inss, bg="white", font=fonte_label, fg="#EF4444").pack(side="left", padx=(0, 20))
        self.var_domestica = tk.IntVar()
        tk.Checkbutton(f_chk, text="Trabalhadora Doméstica", variable=self.var_domestica, bg="white", font=fonte_label, fg="#8B5CF6").pack(side="left")

        tk.Label(frame_esq, text="Últ. Dia Trab:", bg="white", font=fonte_label).grid(row=4, column=0, sticky="w", pady=(0, 8))
        self.ent_ult_dia = self.criar_calendario(frame_esq, width=15)
        self.ent_ult_dia.grid(row=4, column=1, sticky="w", pady=(0, 8), padx=5)

        self.frame_mat = tk.Frame(frame_esq, bg="#FFF1F2", bd=1, relief="solid", padx=10, pady=8)
        self.var_aborto = tk.IntVar()
        tk.Checkbutton(self.frame_mat, text="Afastamento por Aborto", variable=self.var_aborto, bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=0, column=0, sticky="w", pady=4)
        tk.Label(self.frame_mat, text="Data Nascimento:", bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=0, column=1, sticky="e", padx=(10, 5), pady=4)
        self.ent_data_nasc = self.criar_calendario(self.frame_mat)
        self.ent_data_nasc.grid(row=0, column=2, sticky="w", pady=4)

        tk.Label(self.frame_mat, text="Salário Família?:", bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=1, column=0, sticky="w", pady=4)
        self.combo_sal_fam = ttk.Combobox(self.frame_mat, values=["", "Sim", "Não"], width=10, state="readonly", font=fonte_entry)
        self.combo_sal_fam.grid(row=1, column=1, sticky="w", padx=5, pady=4)
        tk.Label(self.frame_mat, text="Obs Sal. Família:", bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=1, column=2, sticky="w", pady=4)
        self.ent_obs_sal_fam = tk.Entry(self.frame_mat, width=25, font=fonte_entry, relief="solid", bd=1)
        self.ent_obs_sal_fam.grid(row=1, column=3, sticky="w", padx=5, pady=4)

        tk.Label(self.frame_mat, text="Imposto Renda?:", bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=2, column=0, sticky="w", pady=4)
        self.combo_imp_renda = ttk.Combobox(self.frame_mat, values=["", "Sim", "Não"], width=10, state="readonly", font=fonte_entry)
        self.combo_imp_renda.grid(row=2, column=1, sticky="w", padx=5, pady=4)
        tk.Label(self.frame_mat, text="Obs Imp. Renda:", bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=2, column=2, sticky="w", pady=4)
        self.ent_obs_imp_renda = tk.Entry(self.frame_mat, width=25, font=fonte_entry, relief="solid", bd=1)
        self.ent_obs_imp_renda.grid(row=2, column=3, sticky="w", padx=5, pady=4)

        def toggle_maternidade(e=None):
            if self.combo_motivo.get() == "Maternidade":
                self.frame_mat.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(4, 4))
            else:
                self.frame_mat.grid_remove()

        self.combo_motivo.bind("<<ComboboxSelected>>", toggle_maternidade)
        self.toggle_maternidade = toggle_maternidade 

        self.frame_aq_box = tk.Frame(frame_esq, bg="#F0F9FF", bd=1, relief="solid", padx=10, pady=8)
        self.frame_aq_box.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(4, 8))

        self.var_ctrl_ferias = tk.IntVar(value=0)
        tk.Checkbutton(
            self.frame_aq_box, text="🏖️ Controlar Período Aquisitivo (Férias - Soma apenas Benefícios Concedidos)",
            variable=self.var_ctrl_ferias, command=self.atualizar_termometro_ferias,
            bg="#F0F9FF", font=("Segoe UI", 9, "bold"), fg="#0369A1"
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))

        tk.Label(self.frame_aq_box, text="Início Per. Aquisitivo:", bg="#F0F9FF", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w")
        self.ent_aq_de = self.criar_calendario(self.frame_aq_box, width=12, cor_fundo='#0284C7')
        self.ent_aq_de.grid(row=1, column=1, sticky="w", padx=(5, 15))
        self.ent_aq_de.bind("<<DateEntrySelected>>", self.atualizar_termometro_ferias)
        self.ent_aq_de.bind("<KeyRelease>", self.atualizar_termometro_ferias)

        tk.Label(self.frame_aq_box, text="Fim Per. Aquisitivo (Auto):", bg="#F0F9FF", font=("Segoe UI", 9, "bold")).grid(row=1, column=2, sticky="w")
        self.ent_aq_ate = tk.Entry(self.frame_aq_box, width=12, font=fonte_entry, relief="solid", bd=1, state="readonly", justify="center")
        self.ent_aq_ate.grid(row=1, column=3, sticky="w", padx=5)

        tk.Label(frame_esq, text="Observações Gerais:", bg="white", font=fonte_label).grid(row=7, column=0, columnspan=4, sticky="w", pady=(4, 2))
        self.txt_obs = tk.Text(frame_esq, width=65, height=4, font=fonte_entry, relief="solid", bd=1)
        self.txt_obs.grid(row=8, column=0, columnspan=4, sticky="w")

        frame_box = ttk.LabelFrame(frame_dir, text=" Período Total do Afastamento ", padding=12)
        frame_box.pack(fill="x", pady=(0, 10))
        
        tk.Label(frame_box, text="Data Inicial (Auto):", bg="white", font=fonte_label).pack(anchor="w", pady=(0, 2))
        self.ent_per_de = tk.Entry(frame_box, font=("Segoe UI", 11), bg="#F3F4F6", relief="solid", bd=1, state="readonly")
        self.ent_per_de.pack(fill="x", pady=(0, 8))

        tk.Label(frame_box, text="Data Final Prevista (Auto):", bg="white", font=fonte_label).pack(anchor="w", pady=(0, 2))
        self.ent_per_ate = tk.Entry(frame_box, font=("Segoe UI", 11), bg="#F3F4F6", relief="solid", bd=1, state="readonly")
        self.ent_per_ate.pack(fill="x", pady=(0, 8))

        tk.Label(frame_box, text="Total de Dias (Período):", bg="white", font=fonte_label).pack(anchor="w", pady=(0, 2))
        self.ent_dias = tk.Entry(frame_box, font=("Segoe UI", 11, "bold"), justify="center", bg="#F3F4F6", relief="solid", bd=1, state="readonly")
        self.ent_dias.pack(fill="x", pady=(0, 8))

        tk.Label(frame_box, text="Soma de Todos os Atestados:", bg="white", font=fonte_label, fg="#0284C7").pack(anchor="w", pady=(0, 2))
        self.ent_soma_atestados = tk.Entry(frame_box, font=("Segoe UI", 11, "bold"), justify="center", bg="#E0F2FE", fg="#0284C7", relief="solid", bd=1, state="readonly")
        self.ent_soma_atestados.pack(fill="x", pady=(0, 12))

        tk.Button(frame_box, text="➕ Gerir Todos os Atestados", command=self.abrir_historico_atestados, bg="#3B82F6", fg="white", font=("Segoe UI", 9, "bold"), pady=6, relief="flat", cursor="hand2").pack(fill="x")

        box_term = ttk.LabelFrame(frame_dir, text=" 🏖️ Termômetro de Férias (Art. 133 CLT) ", padding=12)
        box_term.pack(fill="both", expand=True)

        self.lbl_termometro_dias = tk.Label(box_term, text="0 / 180 dias concedidos", font=("Segoe UI", 13, "bold"), bg="#F3F4F6", fg="#374151", pady=8, relief="solid", bd=1, cursor="hand2")
        self.lbl_termometro_dias.pack(fill="x", pady=(0, 8))
        self.lbl_termometro_dias.bind("<Button-1>", self.abrir_detalhes_ferias)

        self.lbl_termometro_status = tk.Label(box_term, text="", font=("Segoe UI", 9, "bold"), bg="white", wraplength=340, justify="center", cursor="hand2")
        self.lbl_termometro_status.pack(fill="x", pady=(0, 8))
        self.lbl_termometro_status.bind("<Button-1>", self.abrir_detalhes_ferias)

        tk.Button(box_term, text="🔍 Ver Extrato de Dias", relief="flat", fg="#0284C7", bg="white", cursor="hand2", font=("Segoe UI", 9, "underline"), command=self.abrir_detalhes_ferias).pack()

    def toggle_aba_inss(self):
        if self.var_aposentado.get() == 1: self.notebook.tab(self.aba_concessao, state="disabled")
        else: self.notebook.tab(self.aba_concessao, state="normal")

    def atualizar_cid_display(self):
        conn = conectar_db()
        c = conn.cursor()
        c.execute("SELECT codigo FROM cids_historico WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
        cids_res = c.fetchall()
        conn.close()
        self.ent_cid_show.config(state="normal")
        self.ent_cid_show.delete(0, tk.END)
        self.ent_cid_show.insert(0, " | ".join([r[0] for r in cids_res]) if cids_res else "Nenhum")
        self.ent_cid_show.config(state="readonly")

    def abrir_historico_cids(self):
        jan_cid = tk.Toplevel(self.janela)
        jan_cid.title("Gerir CIDs Relacionados")
        jan_cid.geometry("680x420")
        jan_cid.grab_set()
        jan_cid.configure(bg="white")

        tree_cid = ttk.Treeview(jan_cid, columns=("ID", "Código", "Descrição"), show="headings", height=6)
        tree_cid.heading("ID", text="ID")
        tree_cid.heading("Código", text="Cód CID")
        tree_cid.heading("Descrição", text="Descrição / Doença")
        tree_cid.column("ID", width=35, anchor="center")
        tree_cid.column("Código", width=85, anchor="center")
        tree_cid.column("Descrição", width=520, anchor="w")
        tree_cid.pack(fill="x", padx=15, pady=15)

        def carregar_cids():
            for i in tree_cid.get_children(): tree_cid.delete(i)
            conn = conectar_db()
            cursor = conn.cursor()
            cursor.execute("SELECT id, codigo, descricao FROM cids_historico WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
            for linha in cursor.fetchall(): tree_cid.insert("", "end", values=linha)
            conn.close()
            self.atualizar_cid_display()

        carregar_cids()

        def excluir_cid(event):
            selec = tree_cid.selection()
            if not selec: return
            valores = tree_cid.item(selec[0], "values")
            id_cid = valores[0]
            cod_cid = valores[1]
            if messagebox.askyesno("Excluir", f"Remover o CID '{cod_cid}' deste afastamento?", parent=jan_cid):
                conn = conectar_db()
                cursor = conn.cursor()
                cursor.execute("DELETE FROM cids_historico WHERE id=?", (id_cid,))
                cursor.execute("SELECT codigo FROM cids_historico WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
                todos = [r[0] for r in cursor.fetchall()]
                cursor.execute("UPDATE afastamentos SET cid=? WHERE id=?", (" | ".join(todos), self.id_registro))
                conn.commit()
                conn.close()
                carregar_cids()
                self.callback_atualizar()

        tree_cid.bind("<Shift-Delete>", excluir_cid)

        def abrir_modal_editar_cid(event):
            selec = tree_cid.selection()
            if not selec: return
            id_cid, cod_atual, desc_atual = tree_cid.item(selec[0], "values")[:3]

            jan_ed = tk.Toplevel(jan_cid)
            jan_ed.title(f"Editar CID | {cod_atual}")
            jan_ed.geometry("460x180")
            jan_ed.grab_set()
            jan_ed.configure(bg="#F3F4F6")

            box = ttk.LabelFrame(jan_ed, text=" Alterar CID ", padding=15)
            box.pack(fill="both", expand=True, padx=15, pady=15)

            tk.Label(box, text="Código:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
            ed_cod = tk.Entry(box, width=10, font=("Segoe UI", 9), relief="solid", bd=1)
            ed_cod.insert(0, cod_atual)
            ed_cod.grid(row=0, column=1, sticky="w", padx=5, pady=5)

            tk.Label(box, text="Descrição:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=5)
            ed_desc = tk.Entry(box, width=35, font=("Segoe UI", 9), relief="solid", bd=1)
            ed_desc.insert(0, desc_atual)
            ed_desc.grid(row=1, column=1, sticky="w", padx=5, pady=5)

            def salvar_ed_cid():
                c_novo = ed_cod.get().strip().upper()
                d_novo = ed_desc.get().strip()
                if not c_novo:
                    messagebox.showwarning("Aviso", "O código não pode ficar vazio.", parent=jan_ed)
                    return

                c_norm = c_novo.replace(".", "")
                conn = conectar_db()
                cursor = conn.cursor()
                cursor.execute("SELECT codigo FROM cids_historico WHERE id_afastamento=? AND id!=?", (self.id_registro, id_cid))
                existentes = [r[0].strip().upper().replace(".", "") for r in cursor.fetchall()]
                if c_norm in existentes:
                    conn.close()
                    messagebox.showwarning("Duplicidade", f"O CID '{c_novo}' já está cadastrado para este afastamento!", parent=jan_ed)
                    return

                cursor.execute("UPDATE cids_historico SET codigo=?, descricao=? WHERE id=?", (c_novo, d_novo, id_cid))
                cursor.execute("SELECT codigo FROM cids_historico WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
                todos = [r[0] for r in cursor.fetchall()]
                cursor.execute("UPDATE afastamentos SET cid=? WHERE id=?", (" | ".join(todos), self.id_registro))
                conn.commit()
                conn.close()

                carregar_cids()
                self.callback_atualizar()
                jan_ed.destroy()

            tk.Button(jan_ed, text="💾 Salvar Alteração", command=salvar_ed_cid, bg="#3B82F6", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", pady=4, padx=15).pack(pady=(0, 10))

        tree_cid.bind("<Double-1>", abrir_modal_editar_cid)

        frame_add = ttk.LabelFrame(jan_cid, text=" Lançar Novo CID ", padding=15)
        frame_add.pack(fill="both", expand=True, padx=15, pady=5)

        tk.Label(frame_add, text="Código:", bg="white", font=("Segoe UI", 9, "bold")).pack(side="left")
        ent_cod_add = tk.Entry(frame_add, width=9, font=("Segoe UI", 9), relief="solid", bd=1)
        ent_cod_add.pack(side="left", padx=4)

        ent_desc_add = tk.Entry(frame_add, width=38, font=("Segoe UI", 9), relief="solid", bd=1)

        def busca_cid_modal():
            c = ent_cod_add.get().strip().upper()
            if not c: return
            ent_cod_add.delete(0, tk.END)
            ent_cod_add.insert(0, c)
            resultado = buscar_cid_base(c)
            ent_desc_add.delete(0, tk.END)
            if resultado:
                ent_desc_add.insert(0, resultado)
            else:
                ent_desc_add.focus_set()

        ent_cod_add.bind("<Return>", lambda e: busca_cid_modal())

        tk.Button(frame_add, text="🔍 Buscar", command=busca_cid_modal, bg="#3B82F6", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=(0, 10))
        
        tk.Label(frame_add, text="Descrição:", bg="white", font=("Segoe UI", 9, "bold")).pack(side="left")
        ent_desc_add.pack(side="left", padx=4)

        def salvar_cid():
            c = ent_cod_add.get().strip().upper()
            d = ent_desc_add.get().strip()
            if not c: return messagebox.showwarning("Aviso", "Preencha o Código do CID.", parent=jan_cid)

            c_norm = c.replace(".", "")
            conn = conectar_db()
            cursor = conn.cursor()

            cursor.execute("SELECT codigo FROM cids_historico WHERE id_afastamento=?", (self.id_registro,))
            existentes = [r[0].strip().upper().replace(".", "") for r in cursor.fetchall()]
            if c_norm in existentes:
                conn.close()
                messagebox.showwarning("Duplicidade", f"O CID '{c}' já está incluso neste afastamento!", parent=jan_cid)
                return

            cursor.execute("INSERT INTO cids_historico (id_afastamento, codigo, descricao) VALUES (?, ?, ?)", (self.id_registro, c, d))
            cursor.execute("SELECT codigo FROM cids_historico WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
            todos = [r[0] for r in cursor.fetchall()]
            cursor.execute("UPDATE afastamentos SET cid=? WHERE id=?", (" | ".join(todos), self.id_registro))
            conn.commit()
            conn.close()

            ent_cod_add.delete(0, tk.END)
            ent_desc_add.delete(0, tk.END)
            carregar_cids()
            self.callback_atualizar()

        tk.Button(frame_add, text="✔ Adicionar", command=salvar_cid, bg="#10B981", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=10)

        # ================= ABA 2: ACOMPANHAMENTO =================
    def construir_aba_acompanhamento(self):
        f = self.aba_inss
        fonte_label = ("Segoe UI", 10, "bold")
        
        box_lembretes = ttk.LabelFrame(f, text=" Central de Lembretes do Sistema (Popups Matinais) ", padding=15)
        box_lembretes.pack(fill="x", padx=20, pady=(20, 10))

        tk.Label(box_lembretes, text="⚠️ ATENÇÃO: SE ESTIVER MARCADO, O LEMBRETE NÃO ESTARÁ DISPONÍVEL (FICARÁ SILENCIADO)", font=("Segoe UI", 8, "bold"), fg="#DC2626", bg="white").pack(anchor="w", pady=(0, 10))

        self.var_alerta_16 = tk.IntVar()
        tk.Checkbutton(box_lembretes, text="✔ Ocultar Lembrete de Checar Análise Documental (16º Dia)", variable=self.var_alerta_16, bg="white", font=fonte_label, fg="#D97706").pack(anchor="w", pady=2)

        self.var_alerta_mat = tk.IntVar()
        tk.Checkbutton(box_lembretes, text="✔ Ocultar Lembrete de Fim de Licença Maternidade", variable=self.var_alerta_mat, bg="white", font=fonte_label, fg="#7E22CE").pack(anchor="w", pady=2)

        self.var_alerta_retorno = tk.IntVar()
        tk.Checkbutton(box_lembretes, text="✔ Ocultar Lembrete de Verificação de Retorno ao Trabalho", variable=self.var_alerta_retorno, bg="white", font=fonte_label, fg="#166534").pack(anchor="w", pady=2)

        self.var_alerta_ferias = tk.IntVar()
        tk.Checkbutton(box_lembretes, text="✔ Ocultar Lembrete de Perda de Férias (180 Dias de Benefício Concedido)", variable=self.var_alerta_ferias, bg="white", font=fonte_label, fg="#DC2626").pack(anchor="w", pady=2)

        box_agend = ttk.LabelFrame(f, text=" Prazos Iniciais ", padding=30)
        box_agend.pack(fill="x", padx=20, pady=(10, 20))

        tk.Label(box_agend, text="Data da Obrigação:", bg="white", font=fonte_label).grid(row=0, column=0, sticky="w")
        self.ent_obrigacao = self.criar_calendario(box_agend)
        self.ent_obrigacao.grid(row=0, column=1, padx=20, sticky="w")

        self.var_obrig_concluida = tk.IntVar()
        tk.Checkbutton(box_agend, text=" Marcar Obrigação como Tratada", variable=self.var_obrig_concluida, bg="white", font=fonte_label, fg="#10B981").grid(row=0, column=2, padx=30)

        tk.Label(box_agend, text="16º Dia do Afastamento:", bg="white", font=fonte_label).grid(row=1, column=0, sticky="w", pady=(25,0))
        self.ent_16dia = self.criar_calendario(box_agend, cor_fundo='#F59E0B')
        self.ent_16dia.grid(row=1, column=1, padx=20, pady=(25,0), sticky="w")

    # ================= ABA 3: INSS =================
    def gerar_termo_govbr(self):
        nome = self.ent_nome.get().strip()
        cpf = self.ent_cpf.get().strip()
        if not nome or not cpf:
            return messagebox.showwarning("Aviso", "Preencha o Nome e o CPF do funcionário para emitir o termo.", parent=self.janela)
        arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"Termo_GOVBR_{nome.replace(' ', '_')}.pdf", parent=self.janela)
        if not arquivo: return
        try:
            doc = SimpleDocTemplate(arquivo, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
            elementos = []
            estilos = getSampleStyleSheet()
            estilo_titulo = ParagraphStyle('Titulo', parent=estilos['Heading1'], fontSize=12, alignment=TA_CENTER, spaceAfter=20, fontName="Helvetica-Bold")
            estilo_texto = ParagraphStyle('Texto', parent=estilos['Normal'], fontSize=10, alignment=TA_JUSTIFY, spaceAfter=10, leading=14)
            estilo_bold = ParagraphStyle('Negrito', parent=estilos['Normal'], fontSize=10, alignment=TA_LEFT, spaceAfter=8, fontName="Helvetica-Bold")
            estilo_italico = ParagraphStyle('Italico', parent=estilos['Normal'], fontSize=10, alignment=TA_JUSTIFY, fontName="Helvetica-Oblique")
            estilo_linhas = ParagraphStyle('Linhas', parent=estilos['Normal'], fontSize=10, alignment=TA_LEFT, spaceAfter=8)
            
            elementos.append(Paragraph("TERMO DE AUTORIZAÇÃO PARA COLETA, TRATAMENTO E ARMAZENAMENTO DE SENHA GOV.BR", estilo_titulo))
            elementos.append(Paragraph(f"Eu, <b>{nome}</b>, portador do CPF <b>{cpf}</b>,", estilo_texto))
            elementos.append(Paragraph("<b>DECLARO,</b> de forma livre, inequívoca e espontânea, que compartilhei meus dados pessoais, incluindo minha senha de acesso à plataforma GOV.BR à Viman Contabilidade Ltda e autorizo sua utilização para fins de cumprimento de obrigações legais e prestação de serviços previdenciários, tais como requerimentos, perícias e demais procedimentos junto ao INSS.", estilo_texto))
            elementos.append(Paragraph("Estou ciente de que, após verificação de que a finalidade foi alcançada ou de que os dados deixaram de ser necessários ou pertinentes ao alcance da finalidade específica almejada, tenho obrigação de substituir a senha compartilhada, sendo este ato, de minha inteira responsabilidade.", estilo_texto))
            elementos.append(Paragraph("Declaro estar ciente, ainda, de que a não assinatura deste documento implica na impossibilidade de atuação do escritório, sendo, neste caso, considerado que optei por conduzir o processo por conta própria.", estilo_texto))
            elementos.append(Paragraph("Nessa hipótese, declaro estar ciente de que assumo integral responsabilidade pelo acompanhamento de todas as movimentações do processo, comprometendo-me a manter o empregador devidamente informado, mediante o envio de cópias de requerimentos, decisões, exigências e protocolos, bem como pelo cumprimento dos prazos e pela correta apresentação da documentação exigida pelo INSS, ciente de que o descumprimento dessas obrigações poderá resultar no indeferimento do benefício.", estilo_texto))
            elementos.append(Paragraph("Por ser verdade, assino esse documento autorizando, portanto, que os dados informados sejam coletados, tratados e armazenados nos termos deste informativo, sendo certo que quaisquer outros dados que não os já especificados aqui poderão ser solicitados pela Viman Contabilidade Ltda.", estilo_texto))
            elementos.append(Spacer(1, 10))
            
            dados_contatos = [
                [Paragraph("<b>DADOS DE ACESSO AO GOV.BR</b>", estilo_bold), Paragraph("<b>CONTATOS</b>", estilo_bold)],
                [Paragraph(f"Login (CPF): <u>{cpf.ljust(35, '_')}</u>", estilo_linhas), Paragraph("Telefone: " + "_"*35, estilo_linhas)],
                [Paragraph("Senha: " + "_"*42, estilo_linhas), Paragraph("E-mail: " + "_"*38, estilo_linhas)]
            ]
            t1 = Table(dados_contatos, colWidths=[260, 260])
            t1.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0)]))
            elementos.append(t1)
            elementos.append(Spacer(1, 15))
            elementos.append(Paragraph("<b>INFORMAÇÕES PARA REQUERIMENTO DE BENEFÍCIO POR INCAPACIDADE</b>", estilo_bold))
            elementos.append(Paragraph("Data de início dos sintomas: _______/_______/___________", estilo_linhas))
            elementos.append(Paragraph("Descrição dos sintomas: " + "_"*62, estilo_linhas))
            elementos.append(Paragraph("_"*85, estilo_linhas))
            elementos.append(Paragraph("_"*85, estilo_linhas))
            elementos.append(Spacer(1, 10))
            elementos.append(Paragraph("<b>DESCREVA AS ATIVIDADES QUE REALIZA NO TRABALHO (descreva de forma clara e objetiva a rotina de trabalho)</b>", estilo_bold))
            elementos.append(Paragraph("_"*85, estilo_linhas))
            elementos.append(Paragraph("_"*85, estilo_linhas))
            elementos.append(Paragraph("_"*85, estilo_linhas))
            elementos.append(Spacer(1, 30))
            
            assinaturas = [
                ["_"*45, "_"*45],
                [Paragraph("<i>Local / Data</i>", estilo_italico), Paragraph("<i>Assinatura do Cliente / Titular</i>", ParagraphStyle('c', alignment=TA_CENTER, fontName="Helvetica-Oblique", fontSize=10))]
            ]
            t2 = Table(assinaturas, colWidths=[260, 260])
            t2.setStyle(TableStyle([('ALIGN', (0,0), (0,-1), 'LEFT'), ('ALIGN', (1,0), (1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'TOP'), ('BOTTOMPADDING', (0,0), (-1,-1), 2)]))
            elementos.append(t2)
            doc.build(elementos)
            if platform.system() == 'Windows': os.startfile(arquivo)
            elif platform.system() == 'Darwin': subprocess.call(('open', arquivo))
            else: subprocess.call(('xdg-open', arquivo))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar o PDF:\n{str(e)}", parent=self.janela)

    def construir_aba_inss(self):
        f = self.aba_concessao
        fonte_label = ("Segoe UI", 9, "bold")
        fonte_entry = ("Segoe UI", 9)

        box_gov = ttk.LabelFrame(f, text=" Autorização de Acesso (GOV.BR) ", padding=10)
        box_gov.pack(fill="x", padx=20, pady=(12, 5))

        tk.Label(box_gov, text="Tem Autorização no GOV.BR?:", bg="white", font=fonte_label).pack(side="left", padx=5)
        self.combo_gov = ttk.Combobox(box_gov, values=["", "Sim", "Não"], width=10, state="readonly", font=fonte_entry)
        self.combo_gov.pack(side="left", padx=5)
        
        tk.Label(box_gov, text="CPF:", bg="white", font=fonte_label).pack(side="left", padx=(20, 5))
        self.ent_cpf = tk.Entry(box_gov, width=16, font=fonte_entry, relief="solid", bd=1)
        self.ent_cpf.pack(side="left", padx=5)
        self.ent_cpf.bind("<KeyRelease>", self.formatar_cpf)

        tk.Label(box_gov, text="Senha de Acesso:", bg="white", font=fonte_label).pack(side="left", padx=(20, 5))
        self.ent_senha_gov = tk.Entry(box_gov, width=20, font=fonte_entry, relief="solid", bd=1, state="disabled", show="*")
        self.ent_senha_gov.pack(side="left", padx=5)

        tk.Button(box_gov, text="👁️", command=self.pedir_senha_mestre, bg="white", relief="flat", cursor="hand2").pack(side="left")
        tk.Button(box_gov, text="📄 Emitir Autorização", command=self.gerar_termo_govbr, bg="#3B82F6", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=(10, 0))

        self.f_inss_ferias = tk.Frame(f, bg="#F3F4F6", bd=1, relief="solid")
        self.f_inss_ferias.pack(fill="x", padx=20, pady=(0, 5))
        
        self.lbl_inss_ferias = tk.Label(self.f_inss_ferias, text="🏖️ Controle de Férias Desativado", bg="#F3F4F6", font=("Segoe UI", 10, "bold"), fg="#6B7280", pady=6)
        self.lbl_inss_ferias.pack(anchor="center")
        
        def toggle_senha(event=None):
            if self.combo_gov.get() == "Sim": 
                self.ent_senha_gov.config(state="normal")
            else:
                self.ent_senha_gov.config(state="normal")
                self.ent_senha_gov.delete(0, tk.END)
                self.ent_senha_gov.config(state="disabled", show="*")

        self.combo_gov.bind("<<ComboboxSelected>>", toggle_senha)

        box_an = ttk.LabelFrame(f, text=" Histórico de Análises Documentais (INSS) ", padding=10)
        box_an.pack(fill="x", padx=20, pady=5)

        self.tree_an = ttk.Treeview(box_an, columns=("ID", "Data", "Gerou Per?", "Nº Benefício", "Situação", "De", "Até", "Dias", "Prorrog?"), show="headings", height=3)
        for col, w in zip(["ID", "Data", "Gerou Per?", "Nº Benefício", "Situação", "De", "Até", "Dias", "Prorrog?"], [30, 80, 80, 110, 100, 80, 80, 60, 70]):
            self.tree_an.heading(col, text=col)
            self.tree_an.column(col, width=w, anchor="center" if col != "Nº Benefício" else "w")
        self.tree_an.pack(fill="x", pady=(0, 5))
        
        self.tree_an.bind("<Shift-Delete>", self.excluir_analise)
        self.tree_an.bind("<Double-1>", self.abrir_modal_editar_analise)

        f_botoes_tabela_an = tk.Frame(box_an, bg="white")
        f_botoes_tabela_an.pack(fill="x", pady=(0, 10))
        tk.Button(f_botoes_tabela_an, text="🔄 Criar Atestado a partir deste Benefício (Sincronizar)", command=lambda: self.sincronizar_com_beneficio('analise'), bg="#0284C7", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2", padx=10).pack(side="left")
        
        self.lbl_total_an = tk.Label(f_botoes_tabela_an, text="Total de Dias Concedidos: 0", bg="white", font=("Segoe UI", 10, "bold"), fg="#D97706")
        self.lbl_total_an.pack(side="right", padx=15)

        f_an1 = tk.Frame(box_an, bg="white")
        f_an1.pack(fill="x", pady=2)
        tk.Label(f_an1, text="Data Análise:", bg="white", font=fonte_label).pack(side="left")
        self.ent_an_data = self.criar_calendario(f_an1, width=12)
        self.ent_an_data.pack(side="left", padx=(5, 15))
        tk.Label(f_an1, text="Nº Benefício:", bg="white", font=fonte_label).pack(side="left")
        self.ent_an_ben = tk.Entry(f_an1, width=20, font=fonte_entry, relief="solid", bd=1)
        self.ent_an_ben.pack(side="left", padx=(5, 15))
        tk.Label(f_an1, text="Gerou Perícia?:", bg="white", font=fonte_label).pack(side="left")
        self.combo_an_gerou_per = ttk.Combobox(f_an1, values=["", "Sim", "Não"], width=8, state="readonly", font=fonte_entry)
        self.combo_an_gerou_per.pack(side="left", padx=(5, 15))
        tk.Label(f_an1, text="Situação:", bg="white", font=fonte_label).pack(side="left")
        self.combo_an_sit = ttk.Combobox(f_an1, values=["", "Concedido", "Negado", "Em Análise"], width=12, state="readonly", font=fonte_entry)
        self.combo_an_sit.pack(side="left", padx=(5, 15))

        f_an2 = tk.Frame(box_an, bg="white")
        f_an2.pack(fill="x", pady=(5, 0))
        tk.Label(f_an2, text="Concedido De:", bg="white", font=fonte_label).pack(side="left")
        self.ent_an_cde = self.criar_calendario(f_an2, width=12)
        self.ent_an_cde.pack(side="left", padx=(5, 15))
        tk.Label(f_an2, text="Até:", bg="white", font=fonte_label).pack(side="left")
        self.ent_an_cate = self.criar_calendario(f_an2, width=12)
        self.ent_an_cate.pack(side="left", padx=(5, 15))

        def toggle_analise_datas(event=None):
            if self.combo_an_sit.get() == "Concedido":
                self.ent_an_cde.config(state="normal")
                self.ent_an_cate.config(state="normal")
            else:
                for w in [self.ent_an_cde, self.ent_an_cate]:
                    w.config(state="normal")
                    w.delete(0, tk.END)
                    w.config(state="disabled")

        self.combo_an_sit.bind("<<ComboboxSelected>>", toggle_analise_datas)
        toggle_analise_datas()

        tk.Button(f_an2, text="Lançar Análise", command=self.adicionar_analise, bg="#4CAF50", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=10)

        # ---------------- 2. PERÍCIAS INSS (ATUALIZADO) ----------------
        box_per = ttk.LabelFrame(f, text=" Histórico de Perícias e Concessões INSS ", padding=10)
        box_per.pack(fill="both", expand=True, padx=20, pady=5)

        # Nova coluna "Dt. Anál." adicionada com larguras redimensionadas para caber
        self.tree_per = ttk.Treeview(box_per, columns=("ID", "Análise Mãe", "Dt. Anál.", "Agend.", "Atend.", "Espécie", "Benefício", "Situação", "De", "Até", "Dias"), show="headings", height=4)
        for col, w in zip(["ID", "Análise Mãe", "Dt. Anál.", "Agend.", "Atend.", "Espécie", "Benefício", "Situação", "De", "Até", "Dias"], [30, 80, 75, 75, 75, 60, 100, 80, 75, 75, 50]):
            self.tree_per.heading(col, text=col)
            self.tree_per.column(col, width=w, anchor="center" if col != "Benefício" else "w")
        self.tree_per.pack(fill="x", pady=(0, 5))
        
        self.tree_per.bind("<Shift-Delete>", self.excluir_pericia)
        self.tree_per.bind("<Double-1>", self.abrir_modal_editar_pericia)

        f_botoes_tabela_per = tk.Frame(box_per, bg="white")
        f_botoes_tabela_per.pack(fill="x", pady=(0, 10))
        tk.Button(f_botoes_tabela_per, text="🔄 Criar Atestado a partir deste Benefício (Sincronizar)", command=lambda: self.sincronizar_com_beneficio('pericia'), bg="#0284C7", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2", padx=10).pack(side="left")

        self.lbl_total_per = tk.Label(f_botoes_tabela_per, text="Total de Dias Concedidos: 0", bg="white", font=("Segoe UI", 10, "bold"), fg="#D97706")
        self.lbl_total_per.pack(side="right", padx=15)

        f_p1 = tk.Frame(box_per, bg="white")
        f_p1.pack(fill="x", pady=2)
        tk.Label(f_p1, text="Análise Mãe:", bg="white", font=fonte_label, fg="#EC4899").pack(side="left")
        self.combo_per_mae = ttk.Combobox(f_p1, width=18, state="readonly", font=fonte_entry)
        self.combo_per_mae.pack(side="left", padx=(5, 10))
        
        # NOVO CAMPO: Data de Análise
        tk.Label(f_p1, text="Dt. Análise:", bg="white", font=fonte_label).pack(side="left")
        self.ent_per_dt_analise = self.criar_calendario(f_p1, width=11)
        self.ent_per_dt_analise.pack(side="left", padx=(5, 10))

        tk.Label(f_p1, text="Agendamento:", bg="white", font=fonte_label).pack(side="left")
        self.ent_per_agend = self.criar_calendario(f_p1, width=11)
        self.ent_per_agend.pack(side="left", padx=(5, 10))
        tk.Label(f_p1, text="Atendimento:", bg="white", font=fonte_label).pack(side="left")
        self.ent_per_atend = self.criar_calendario(f_p1, width=11)
        self.ent_per_atend.pack(side="left", padx=(5, 10))
        
        f_p2 = tk.Frame(box_per, bg="white")
        f_p2.pack(fill="x", pady=(5, 0))
        tk.Label(f_p2, text="Espécie:", bg="white", font=fonte_label).pack(side="left")
        self.ent_per_esp = tk.Entry(f_p2, width=10, font=fonte_entry, relief="solid", bd=1)
        self.ent_per_esp.pack(side="left", padx=(5, 15))
        tk.Label(f_p2, text="Nº Benefício:", bg="white", font=fonte_label).pack(side="left")
        self.ent_per_ben = tk.Entry(f_p2, width=20, font=fonte_entry, relief="solid", bd=1)
        self.ent_per_ben.pack(side="left", padx=(5, 15))
        tk.Label(f_p2, text="Situação:", bg="white", font=fonte_label).pack(side="left")
        self.combo_per_sit = ttk.Combobox(f_p2, values=["", "Concedida", "Negada", "Em Análise"], width=12, state="readonly", font=fonte_entry)
        self.combo_per_sit.pack(side="left", padx=(5, 15))
        
        f_p3 = tk.Frame(box_per, bg="white")
        f_p3.pack(fill="x", pady=(5, 0))
        tk.Label(f_p3, text="Concedido De:", bg="white", font=fonte_label).pack(side="left")
        self.ent_per_cde = self.criar_calendario(f_p3, width=12)
        self.ent_per_cde.pack(side="left", padx=(5, 15))
        tk.Label(f_p3, text="Até:", bg="white", font=fonte_label).pack(side="left")
        self.ent_per_cate = self.criar_calendario(f_p3, width=12)
        self.ent_per_cate.pack(side="left", padx=(5, 15))

        tk.Button(f_p3, text="Lançar Perícia", command=self.adicionar_pericia, bg="#4CAF50", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=10)

        self.carregar_listas_inss()

    def formatar_cpf(self, event=None):
        if event and event.keysym in ('BackSpace', 'Delete', 'Left', 'Right'): return
        numeros = "".join(filter(str.isdigit, self.ent_cpf.get()))[:11]
        resultado = ""
        for i, char in enumerate(numeros):
            if i in [3, 6]: resultado += "."
            elif i == 9: resultado += "-"
            resultado += char
        self.ent_cpf.delete(0, tk.END)
        self.ent_cpf.insert(0, resultado)

    def pedir_senha_mestre(self):
        if self.ent_senha_gov.cget('state') == 'disabled': return
        if self.ent_senha_gov.cget('show') == '':
            self.ent_senha_gov.config(show="*")
            return
        pwd = simpledialog.askstring("Segurança", "Digite a senha mestre para visualizar:", parent=self.janela, show='*')
        if pwd == "@Viman1234@":
            self.ent_senha_gov.config(show="")
        elif pwd is not None:
            messagebox.showerror("Acesso Negado", "Senha mestre incorreta.", parent=self.janela)

    # ================= MÉTODOS INSS =================
    def carregar_listas_inss(self):
        for i in self.tree_an.get_children(): self.tree_an.delete(i)
        for i in self.tree_per.get_children(): self.tree_per.delete(i)
        
        conn = conectar_db()
        cursor = conn.cursor()
        
        cursor.execute("SELECT id, data_analise, gerou_pericia, numero_beneficio, situacao, concessao_de, concessao_ate, prorrogacao_tratada FROM analises_inss WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
        analises = cursor.fetchall()
        
        self.map_analises_mae.clear()
        lista_combobox = [""] 
        total_an_dias = 0
        
        for an in analises:
            dias_val = "-"
            if an[4] in ("Concedido", "Concedida") and an[5] and an[6]:
                try:
                    dias = (datetime.strptime(an[6], "%d/%m/%Y") - datetime.strptime(an[5], "%d/%m/%Y")).days + 1
                    if dias > 0:
                        dias_val = str(dias)
                        total_an_dias += dias
                except ValueError: pass

            prorrog_str = "Sim" if an[7] == 1 else "Não"
            self.tree_an.insert("", "end", values=(an[0], an[1], an[2], an[3], an[4], an[5], an[6], dias_val, prorrog_str))
            if an[2] == "Sim":
                lbl = f"Análise #{an[0]} (Data: {an[1]})" 
                self.map_analises_mae[lbl] = an[0]
                lista_combobox.append(lbl)
                
        self.combo_per_mae['values'] = lista_combobox
        self.lbl_total_an.config(text=f"Total de Dias Concedidos: {total_an_dias}")
        
        # Trazendo agora o campo data_analise_pericia
        cursor.execute("SELECT id, id_analise_mae, data_analise_pericia, data_agendamento, data_atendimento, especie, numero_beneficio, situacao, concessao_de, concessao_ate FROM pericias_inss WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
        total_per_dias = 0

        for p in cursor.fetchall():
            dias_val = "-"
            if p[8] in ("Concedido", "Concedida") and p[9] and p[10]:
                try:
                    dias = (datetime.strptime(p[10], "%d/%m/%Y") - datetime.strptime(p[9], "%d/%m/%Y")).days + 1
                    if dias > 0:
                        dias_val = str(dias)
                        total_per_dias += dias
                except ValueError: pass

            mae_texto = f"Análise #{p[1]}" if p[1] else "-"
            # values = (ID, Análise Mãe, Dt. Análise, Agendamento, Atendimento, Espécie, Benefício, Situação, De, Até, Dias)
            self.tree_per.insert("", "end", values=(p[0], mae_texto, p[2], p[3], p[4], p[5], p[6], p[7], p[8], p[9], dias_val))
            
        self.lbl_total_per.config(text=f"Total de Dias Concedidos: {total_per_dias}")
        conn.close()

        if hasattr(self, 'ent_aq_de'):
            self.atualizar_termometro_ferias()

    def adicionar_analise(self):
        dt, n_ben, g_per, sit, cde, cate = self.ent_an_data.get(), self.ent_an_ben.get(), self.combo_an_gerou_per.get(), self.combo_an_sit.get(), self.ent_an_cde.get(), self.ent_an_cate.get()
        if sit in ["Concedido", "Concedida"] and cde and cate:
            try:
                dt_de, dt_ate = datetime.strptime(cde, "%d/%m/%Y"), datetime.strptime(cate, "%d/%m/%Y")
                if dt_ate < dt_de: return messagebox.showwarning("Erro", "A data 'Até' não pode ser anterior a 'De'.", parent=self.janela)
                conn = conectar_db()
                cursor = conn.cursor()
                cursor.execute("SELECT concessao_de, concessao_ate FROM analises_inss WHERE id_afastamento=? AND situacao IN ('Concedido', 'Concedida')", (self.id_registro,))
                for ext in cursor.fetchall():
                    if ext[0] and ext[1]:
                        if dt_de <= datetime.strptime(ext[1], "%d/%m/%Y") and datetime.strptime(ext[0], "%d/%m/%Y") <= dt_ate:
                            conn.close()
                            return messagebox.showwarning("Aviso de Sobreposição", f"A concessão ({cde} a {cate}) sobrepõe-se a uma existente ({ext[0]} a {ext[1]}).", parent=self.janela)
                conn.close()
            except ValueError: return messagebox.showwarning("Erro", "Datas de concessão inválidas.", parent=self.janela)

        conn = conectar_db()
        conn.cursor().execute("INSERT INTO analises_inss (id_afastamento, data_analise, numero_analise, gerou_pericia, gerou_beneficio, numero_beneficio, situacao, concessao_de, concessao_ate, prorrogacao_tratada) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", 
                              (self.id_registro, dt, "", g_per, "", n_ben, sit, cde, cate, 0))
        conn.commit()
        conn.close()
        self.ent_an_data.delete(0, tk.END)
        self.ent_an_ben.delete(0, tk.END)
        self.combo_an_gerou_per.set("")
        self.combo_an_sit.set("")
        self.combo_an_sit.event_generate("<<ComboboxSelected>>") 
        self.carregar_listas_inss()

    def excluir_analise(self, event):
        selec = self.tree_an.selection()
        if not selec: return
        id_item = self.tree_an.item(selec[0], "values")[0]
        if messagebox.askyesno("Excluir", f"Apagar análise ID {id_item} e suas perícias?", parent=self.janela):
            conn = conectar_db()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM pericias_inss WHERE id_analise_mae=?", (id_item,))
            cursor.execute("DELETE FROM analises_inss WHERE id=?", (id_item,))
            conn.commit()
            conn.close()
            self.carregar_listas_inss()

    def adicionar_pericia(self):
        id_mae = self.map_analises_mae.get(self.combo_per_mae.get(), None)
        dt_analise = self.ent_per_dt_analise.get()
        agend, atend, esp, ben, sit, cde, cate = self.ent_per_agend.get(), self.ent_per_atend.get(), self.ent_per_esp.get(), self.ent_per_ben.get(), self.combo_per_sit.get(), self.ent_per_cde.get(), self.ent_per_cate.get()
        if sit in ["Concedido", "Concedida"] and cde and cate:
            try:
                dt_de, dt_ate = datetime.strptime(cde, "%d/%m/%Y"), datetime.strptime(cate, "%d/%m/%Y")
                if dt_ate < dt_de: return messagebox.showwarning("Erro", "A data 'Até' não pode ser anterior a 'De'.", parent=self.janela)
                conn = conectar_db()
                cursor = conn.cursor()
                cursor.execute("SELECT concessao_de, concessao_ate FROM pericias_inss WHERE id_afastamento=? AND situacao IN ('Concedido', 'Concedida')", (self.id_registro,))
                for ext in cursor.fetchall():
                    if ext[0] and ext[1]:
                        if dt_de <= datetime.strptime(ext[1], "%d/%m/%Y") and datetime.strptime(ext[0], "%d/%m/%Y") <= dt_ate:
                            conn.close()
                            return messagebox.showwarning("Aviso de Sobreposição", f"A concessão ({cde} a {cate}) sobrepõe-se a uma existente ({ext[0]} a {ext[1]}).", parent=self.janela)
                conn.close()
            except ValueError: return messagebox.showwarning("Erro", "Datas de concessão inválidas.", parent=self.janela)

        conn = conectar_db()
        conn.cursor().execute("INSERT INTO pericias_inss (id_afastamento, id_analise_mae, data_analise_pericia, data_agendamento, data_atendimento, especie, numero_beneficio, situacao, concessao_de, concessao_ate) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", 
                              (self.id_registro, id_mae, dt_analise, agend, atend, esp, ben, sit, cde, cate))
        conn.commit()
        conn.close()
        self.combo_per_mae.set("")
        for w in [self.ent_per_dt_analise, self.ent_per_agend, self.ent_per_atend, self.ent_per_esp, self.ent_per_ben, self.ent_per_cde, self.ent_per_cate]: w.delete(0, tk.END)
        self.combo_per_sit.set("")
        self.carregar_listas_inss()

    def excluir_pericia(self, event):
        selec = self.tree_per.selection()
        if not selec: return
        id_item = self.tree_per.item(selec[0], "values")[0]
        if messagebox.askyesno("Excluir", f"Apagar perícia ID {id_item}?", parent=self.janela):
            conn = conectar_db()
            conn.cursor().execute("DELETE FROM pericias_inss WHERE id=?", (id_item,))
            conn.commit()
            conn.close()
            self.carregar_listas_inss()

    def abrir_modal_editar_analise(self, event):
        selec = self.tree_an.selection()
        if not selec: return
        id_analise = self.tree_an.item(selec[0], "values")[0]
        conn = conectar_db()
        cursor = conn.cursor()
        cursor.execute("SELECT data_analise, numero_beneficio, gerou_pericia, situacao, concessao_de, concessao_ate, prorrogacao_tratada FROM analises_inss WHERE id=?", (id_analise,))
        dados_an = cursor.fetchone()
        conn.close()
        if not dados_an: return
        
        jan = tk.Toplevel(self.janela)
        jan.title(f"Editar Análise INSS | ID: {id_analise}")
        jan.geometry("520x350")
        jan.grab_set()
        jan.configure(bg="#F3F4F6")
        box = ttk.LabelFrame(jan, text=" Detalhes da Análise ", padding=20)
        box.pack(fill="both", expand=True, padx=20, pady=20)
        for i in range(4): box.columnconfigure(i, weight=1)
        fonte_lbl = ("Segoe UI", 9, "bold")

        tk.Label(box, text="Data Análise:", bg="white", font=fonte_lbl).grid(row=0, column=0, sticky="w", pady=(0, 10))
        ent_dt = self.criar_calendario(box, width=15)
        ent_dt.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=(0, 10))
        tk.Label(box, text="Nº Benefício:", bg="white", font=fonte_lbl).grid(row=0, column=2, sticky="w", pady=(0, 10))
        ent_ben = tk.Entry(box, font=("Segoe UI", 9), relief="solid", bd=1)
        ent_ben.grid(row=0, column=3, sticky="ew", pady=(0, 10))

        tk.Label(box, text="Gerou Perícia?:", bg="white", font=fonte_lbl).grid(row=1, column=0, sticky="w", pady=(0, 10))
        cb_gerou = ttk.Combobox(box, values=["", "Sim", "Não"], state="readonly", font=("Segoe UI", 9))
        cb_gerou.grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=(0, 10))
        tk.Label(box, text="Situação:", bg="white", font=fonte_lbl).grid(row=1, column=2, sticky="w", pady=(0, 10))
        cb_sit = ttk.Combobox(box, values=["", "Concedido", "Negado", "Em Análise"], state="readonly", font=("Segoe UI", 9))
        cb_sit.grid(row=1, column=3, sticky="ew", pady=(0, 10))

        ttk.Separator(box, orient="horizontal").grid(row=2, column=0, columnspan=4, sticky="ew", pady=15)

        tk.Label(box, text="Concedido De:", bg="white", font=fonte_lbl).grid(row=3, column=0, sticky="w", pady=(0, 10))
        ent_cde = self.criar_calendario(box, width=15)
        ent_cde.grid(row=3, column=1, sticky="ew", padx=(0, 10), pady=(0, 10))
        tk.Label(box, text="Até:", bg="white", font=fonte_lbl).grid(row=3, column=2, sticky="w", pady=(0, 10))
        ent_cate = self.criar_calendario(box, width=15)
        ent_cate.grid(row=3, column=3, sticky="ew", pady=(0, 10))

        var_prorrog = tk.IntVar(value=dados_an[6] if dados_an[6] else 0)
        chk_prorrog = tk.Checkbutton(box, text="✔ Prorrogação Checada / Tratada", variable=var_prorrog, bg="white", font=fonte_lbl, fg="#0284C7")
        chk_prorrog.grid(row=4, column=0, columnspan=4, sticky="w", pady=(10, 0))

        def t_datas(e=None):
            if cb_sit.get() == "Concedido":
                ent_cde.config(state="normal")
                ent_cate.config(state="normal")
                chk_prorrog.config(state="normal")
            else:
                for w in [ent_cde, ent_cate]:
                    w.config(state="normal")
                    w.delete(0, tk.END)
                    w.config(state="disabled")
                chk_prorrog.config(state="disabled")
                var_prorrog.set(0)

        cb_sit.bind("<<ComboboxSelected>>", t_datas)
        self.preencher_data_segura(ent_dt, dados_an[0])
        ent_ben.insert(0, dados_an[1] or "")
        cb_gerou.set(dados_an[2] or "")
        cb_sit.set(dados_an[3] or "")
        t_datas()
        self.preencher_data_segura(ent_cde, dados_an[4])
        self.preencher_data_segura(ent_cate, dados_an[5])

        def sav():
            cde_val, cate_val, sit_val = ent_cde.get(), ent_cate.get(), cb_sit.get()
            if sit_val in ["Concedido", "Concedida"] and cde_val and cate_val:
                try:
                    dt_de, dt_ate = datetime.strptime(cde_val, "%d/%m/%Y"), datetime.strptime(cate_val, "%d/%m/%Y")
                    if dt_ate < dt_de: return messagebox.showwarning("Erro", "A data 'Até' não pode ser anterior a 'De'.", parent=jan)
                    conn = conectar_db()
                    cursor = conn.cursor()
                    cursor.execute("SELECT concessao_de, concessao_ate FROM analises_inss WHERE id_afastamento=? AND situacao IN ('Concedido', 'Concedida') AND id!=?", (self.id_registro, id_analise))
                    for ext in cursor.fetchall():
                        if ext[0] and ext[1]:
                            if dt_de <= datetime.strptime(ext[1], "%d/%m/%Y") and datetime.strptime(ext[0], "%d/%m/%Y") <= dt_ate:
                                conn.close()
                                return messagebox.showwarning("Aviso de Sobreposição", f"A concessão ({cde_val} a {cate_val}) sobrepõe-se a uma existente ({ext[0]} a {ext[1]}).", parent=jan)
                    conn.close()
                except ValueError: return messagebox.showwarning("Erro", "Datas inválidas.", parent=jan)

            conn = conectar_db()
            conn.cursor().execute("UPDATE analises_inss SET data_analise=?, numero_beneficio=?, gerou_pericia=?, situacao=?, concessao_de=?, concessao_ate=?, prorrogacao_tratada=? WHERE id=?", 
                                  (ent_dt.get(), ent_ben.get(), cb_gerou.get(), sit_val, cde_val, cate_val, var_prorrog.get(), id_analise))
            conn.commit()
            conn.close()
            self.carregar_listas_inss()
            jan.destroy()

        tk.Button(jan, text="💾 Guardar Alterações", command=sav, bg="#3B82F6", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", pady=5, padx=20).pack(pady=(0, 20))

    def abrir_modal_editar_pericia(self, event):
        selec = self.tree_per.selection()
        if not selec: return
        id_pericia = self.tree_per.item(selec[0], "values")[0]
        conn = conectar_db()
        cursor = conn.cursor()
        cursor.execute("SELECT id_analise_mae, data_analise_pericia, data_agendamento, data_atendimento, especie, numero_beneficio, situacao, concessao_de, concessao_ate FROM pericias_inss WHERE id=?", (id_pericia,))
        dados_p = cursor.fetchone()
        conn.close()
        if not dados_p: return
        
        jan = tk.Toplevel(self.janela)
        jan.title(f"Editar Perícia INSS | ID: {id_pericia}")
        jan.geometry("580x420")
        jan.grab_set()
        jan.configure(bg="#F3F4F6")
        box = ttk.LabelFrame(jan, text=" Detalhes da Perícia ", padding=20)
        box.pack(fill="both", expand=True, padx=20, pady=20)
        for i in range(4): box.columnconfigure(i, weight=1)
        fonte_lbl = ("Segoe UI", 9, "bold")

        tk.Label(box, text="Análise Mãe:", bg="white", font=fonte_lbl, fg="#EC4899").grid(row=0, column=0, sticky="w", pady=(0, 10))
        cb_mae = ttk.Combobox(box, state="readonly", font=("Segoe UI", 9), values=list(self.map_analises_mae.keys()))
        cb_mae.grid(row=0, column=1, columnspan=3, sticky="ew", pady=(0, 10))

        tk.Label(box, text="Data Análise:", bg="white", font=fonte_lbl).grid(row=1, column=0, sticky="w", pady=(0, 10))
        ent_dt_analise = self.criar_calendario(box, width=15)
        ent_dt_analise.grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=(0, 10))

        tk.Label(box, text="Agendamento:", bg="white", font=fonte_lbl).grid(row=1, column=2, sticky="w", pady=(0, 10))
        ent_agend = self.criar_calendario(box, width=15)
        ent_agend.grid(row=1, column=3, sticky="ew", pady=(0, 10))
        
        tk.Label(box, text="Atendimento:", bg="white", font=fonte_lbl).grid(row=2, column=0, sticky="w", pady=(0, 10))
        ent_atend = self.criar_calendario(box, width=15)
        ent_atend.grid(row=2, column=1, sticky="ew", padx=(0, 10), pady=(0, 10))

        tk.Label(box, text="Espécie:", bg="white", font=fonte_lbl).grid(row=2, column=2, sticky="w", pady=(0, 10))
        ent_esp = tk.Entry(box, font=("Segoe UI", 9), relief="solid", bd=1)
        ent_esp.grid(row=2, column=3, sticky="ew", pady=(0, 10))

        tk.Label(box, text="Nº Benefício:", bg="white", font=fonte_lbl).grid(row=3, column=0, sticky="w", pady=(0, 10))
        ent_ben = tk.Entry(box, font=("Segoe UI", 9), relief="solid", bd=1)
        ent_ben.grid(row=3, column=1, sticky="ew", padx=(0, 10), pady=(0, 10))

        tk.Label(box, text="Situação:", bg="white", font=fonte_lbl).grid(row=3, column=2, sticky="w", pady=(0, 10))
        cb_sit = ttk.Combobox(box, values=["", "Concedida", "Negada", "Em Análise"], state="readonly", font=("Segoe UI", 9))
        cb_sit.grid(row=3, column=3, sticky="ew", pady=(0, 10))

        ttk.Separator(box, orient="horizontal").grid(row=4, column=0, columnspan=4, sticky="ew", pady=15)

        tk.Label(box, text="Concedido De:", bg="white", font=fonte_lbl).grid(row=5, column=0, sticky="w", pady=(0, 10))
        ent_cde = self.criar_calendario(box, width=15)
        ent_cde.grid(row=5, column=1, sticky="ew", padx=(0, 10), pady=(0, 10))
        tk.Label(box, text="Até:", bg="white", font=fonte_lbl).grid(row=5, column=2, sticky="w", pady=(0, 10))
        ent_cate = self.criar_calendario(box, width=15)
        ent_cate.grid(row=5, column=3, sticky="ew", pady=(0, 10))

        rev_map = {v: k for k, v in self.map_analises_mae.items()}
        cb_mae.set(rev_map.get(dados_p[0], ""))
        self.preencher_data_segura(ent_dt_analise, dados_p[1])
        self.preencher_data_segura(ent_agend, dados_p[2])
        self.preencher_data_segura(ent_atend, dados_p[3])
        ent_esp.insert(0, dados_p[4] or "")
        ent_ben.insert(0, dados_p[5] or "")
        cb_sit.set(dados_p[6] or "")
        self.preencher_data_segura(ent_cde, dados_p[7])
        self.preencher_data_segura(ent_cate, dados_p[8])

        def sav_per():
            cde_val, cate_val, sit_val = ent_cde.get(), ent_cate.get(), cb_sit.get()
            if sit_val in ["Concedido", "Concedida"] and cde_val and cate_val:
                try:
                    dt_de, dt_ate = datetime.strptime(cde_val, "%d/%m/%Y"), datetime.strptime(cate_val, "%d/%m/%Y")
                    if dt_ate < dt_de: return messagebox.showwarning("Erro", "A data 'Até' não pode ser anterior a 'De'.", parent=jan)
                    conn = conectar_db()
                    cursor = conn.cursor()
                    cursor.execute("SELECT concessao_de, concessao_ate FROM pericias_inss WHERE id_afastamento=? AND situacao IN ('Concedido', 'Concedida') AND id!=?", (self.id_registro, id_pericia))
                    for ext in cursor.fetchall():
                        if ext[0] and ext[1]:
                            if dt_de <= datetime.strptime(ext[1], "%d/%m/%Y") and datetime.strptime(ext[0], "%d/%m/%Y") <= dt_ate:
                                conn.close()
                                return messagebox.showwarning("Aviso de Sobreposição", f"A concessão ({cde_val} a {cate_val}) sobrepõe-se a uma existente ({ext[0]} a {ext[1]}).", parent=jan)
                    conn.close()
                except ValueError: return messagebox.showwarning("Erro", "Datas inválidas.", parent=jan)
            
            id_m = self.map_analises_mae.get(cb_mae.get(), None)
            conn = conectar_db()
            conn.cursor().execute("UPDATE pericias_inss SET id_analise_mae=?, data_analise_pericia=?, data_agendamento=?, data_atendimento=?, especie=?, numero_beneficio=?, situacao=?, concessao_de=?, concessao_ate=? WHERE id=?", 
                                  (id_m, ent_dt_analise.get(), ent_agend.get(), ent_atend.get(), ent_esp.get(), ent_ben.get(), sit_val, cde_val, cate_val, id_pericia))
            conn.commit()
            conn.close()
            self.carregar_listas_inss()
            jan.destroy()

        tk.Button(jan, text="💾 Guardar Alterações", command=sav_per, bg="#3B82F6", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", pady=5, padx=20).pack(pady=(0, 20))

    def abrir_modal_recurso(self):
        TelaRecurso(self.janela, self.id_registro, self.ent_empresa.get(), self.ent_nome.get())

    # ================= ABA 4: RETORNO E NOVO PERÍODO AQUISITIVO =================
    def construir_aba_retorno(self):
        f = self.aba_retorno
        fonte_label = ("Segoe UI", 10, "bold")

        box_retorno = ttk.LabelFrame(f, text=" Fecho do Afastamento ", padding=30)
        box_retorno.pack(fill="x", padx=20, pady=(20, 10))
        
        self.var_retornou = tk.IntVar()
        chk = tk.Checkbutton(box_retorno, text=" O funcionário já retornou definitivamente ao trabalho?", variable=self.var_retornou, command=self.verificar_virada_aquisitivo_retorno, bg="white", font=("Segoe UI", 12, "bold"), fg="#10B981")
        chk.pack(anchor="w", pady=(0, 25))

        frame_campos = tk.Frame(box_retorno, bg="white")
        frame_campos.pack(fill="x")

        tk.Label(frame_campos, text="Data Efetiva de Retorno:", bg="white", font=fonte_label).grid(row=0, column=0, sticky="w", pady=(0, 15))
        self.ent_ret_data = self.criar_calendario(frame_campos, cor_fundo='#10B981')
        self.ent_ret_data.grid(row=0, column=1, padx=15, pady=(0, 15), sticky="w")
        self.ent_ret_data.bind("<<DateEntrySelected>>", self.verificar_virada_aquisitivo_retorno)
        self.ent_ret_data.bind("<KeyRelease>", self.verificar_virada_aquisitivo_retorno)

        tk.Label(frame_campos, text="Fim da Estabilidade:", bg="white", font=fonte_label).grid(row=0, column=2, sticky="w", padx=(30, 0), pady=(0, 15))
        self.ent_estab = self.criar_calendario(frame_campos)
        self.ent_estab.grid(row=0, column=3, padx=15, pady=(0, 15), sticky="w")

        tk.Label(frame_campos, text="HISCRE Tirado? (Obrigatório em Acidentes):", bg="white", font=fonte_label).grid(row=1, column=0, columnspan=2, sticky="w")
        self.combo_hiscre = ttk.Combobox(frame_campos, values=["", "Sim", "Não"], width=10, state="readonly", font=("Segoe UI", 10))
        self.combo_hiscre.grid(row=1, column=2, padx=15, sticky="w")

        box_ferias_ret = ttk.LabelFrame(f, text=" 🏖️ Impacto nas Férias e Novo Período Aquisitivo (Art. 133, § 2º CLT) ", padding=25)
        box_ferias_ret.pack(fill="both", expand=True, padx=20, pady=10)

        self.lbl_aviso_perda_retorno = tk.Label(box_ferias_ret, text="", bg="white", font=("Segoe UI", 10, "bold"), justify="left")
        self.lbl_aviso_perda_retorno.pack(anchor="w", pady=(0, 15))

        f_novo_aq = tk.Frame(box_ferias_ret, bg="white")
        f_novo_aq.pack(fill="x")

        tk.Label(f_novo_aq, text="Novo Período Aquisitivo (Início no Retorno):", bg="white", font=fonte_label, fg="#0369A1").grid(row=0, column=0, sticky="w")
        self.ent_novo_aq_de = self.criar_calendario(f_novo_aq, width=14, cor_fundo='#0284C7')
        self.ent_novo_aq_de.grid(row=0, column=1, padx=(10, 25), sticky="w")
        self.ent_novo_aq_de.bind("<<DateEntrySelected>>", self.atualizar_fim_novo_aquisitivo)
        self.ent_novo_aq_de.bind("<KeyRelease>", self.atualizar_fim_novo_aquisitivo)

        tk.Label(f_novo_aq, text="Fim do Novo Período (Auto):", bg="white", font=fonte_label, fg="#0369A1").grid(row=0, column=2, sticky="w")
        self.ent_novo_aq_ate = tk.Entry(f_novo_aq, width=14, font=("Segoe UI", 10, "bold"), relief="solid", bd=1, state="readonly", justify="center", bg="#F3F4F6")
        self.ent_novo_aq_ate.grid(row=0, column=3, padx=10, sticky="w")

        self.btn_termo_ferias = tk.Button(box_ferias_ret, text="📄 Emitir Termo de Perda de Férias", command=self.gerar_termo_perda_ferias, bg="#DC2626", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", padx=15, pady=6)

    def gerar_termo_perda_ferias(self):
        emitir_termo_ferias(self)

    # ================= HISTÓRICO DE ATESTADOS =================
    def abrir_historico_atestados(self):
        jan_hist = tk.Toplevel(self.janela)
        jan_hist.title("Gerir Atestados e Prorrogações")
        jan_hist.geometry("700x450")
        jan_hist.grab_set()
        jan_hist.configure(bg="white")

        tree_hist = ttk.Treeview(jan_hist, columns=("ID", "De", "Até", "Dias", "Obs"), show="headings", height=6)
        for col, text, w in [("ID", "ID", 30), ("De", "Data De", 80), ("Até", "Data Até", 80), ("Dias", "Dias", 50), ("Obs", "Observações", 250)]:
            tree_hist.heading(col, text=text)
            tree_hist.column(col, width=w, anchor="center" if col != "Obs" else "w")
        tree_hist.pack(fill="x", padx=15, pady=15)

        def carregar_lista_hist():
            for i in tree_hist.get_children(): tree_hist.delete(i)
            conn = conectar_db()
            cursor = conn.cursor()
            cursor.execute("SELECT id, data_de, data_ate, qtd_dias, observacoes FROM atestados_historico WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
            for linha in cursor.fetchall(): tree_hist.insert("", "end", values=linha)
            conn.close()

        carregar_lista_hist()

        def excluir_atestado(event):
            selecionado = tree_hist.selection()
            if not selecionado: return
            id_hist = tree_hist.item(selecionado[0], "values")[0]
            if messagebox.askyesno("Excluir", f"Apagar definitivamente este atestado (ID: {id_hist})?", parent=jan_hist):
                conn = conectar_db()
                conn.cursor().execute("DELETE FROM atestados_historico WHERE id=?", (id_hist,))
                conn.commit()
                conn.close()
                carregar_lista_hist()
                self.sincronizar_totais_afastamento()

        tree_hist.bind("<Shift-Delete>", excluir_atestado)

        def abrir_modal_editar_atestado(event):
            selecionado = tree_hist.selection()
            if not selecionado: return
            id_atestado = tree_hist.item(selecionado[0], "values")[0]
            conn = conectar_db()
            cursor = conn.cursor()
            cursor.execute("SELECT data_de, data_ate, observacoes FROM atestados_historico WHERE id=?", (id_atestado,))
            dados_atest = cursor.fetchone()
            conn.close()
            if not dados_atest: return
            
            jan_ed = tk.Toplevel(jan_hist)
            jan_ed.title(f"Editar Atestado | ID: {id_atestado}")
            jan_ed.geometry("450x250")
            jan_ed.grab_set()
            jan_ed.configure(bg="#F3F4F6")
            box = ttk.LabelFrame(jan_ed, text=" Modificar Atestado ", padding=20)
            box.pack(fill="both", expand=True, padx=20, pady=20)
            fonte_lbl = ("Segoe UI", 9, "bold")
            
            tk.Label(box, text="De:", bg="white", font=fonte_lbl).grid(row=0, column=0, sticky="w", pady=5)
            ed_de = self.criar_calendario(box, width=12)
            ed_de.grid(row=0, column=1, sticky="w", padx=10, pady=5)
            tk.Label(box, text="Até:", bg="white", font=fonte_lbl).grid(row=0, column=2, sticky="w", pady=5)
            ed_ate = self.criar_calendario(box, width=12)
            ed_ate.grid(row=0, column=3, sticky="w", padx=10, pady=5)
            tk.Label(box, text="Obs:", bg="white", font=fonte_lbl).grid(row=1, column=0, sticky="w", pady=15)
            ed_obs = tk.Entry(box, width=35, font=("Segoe UI", 9), relief="solid", bd=1)
            ed_obs.grid(row=1, column=1, columnspan=3, sticky="w", padx=10, pady=15)
            
            self.preencher_data_segura(ed_de, dados_atest[0])
            self.preencher_data_segura(ed_ate, dados_atest[1])
            ed_obs.insert(0, dados_atest[2] or "")
            
            def salvar_ed_atestado():
                v_de, v_ate = ed_de.get(), ed_ate.get()
                try: 
                    dt_de, dt_ate = datetime.strptime(v_de, "%d/%m/%Y"), datetime.strptime(v_ate, "%d/%m/%Y")
                    if dt_ate < dt_de: return messagebox.showwarning("Erro", "A data 'Até' não pode ser anterior a 'De'.", parent=jan_ed)
                    conn = conectar_db()
                    cursor = conn.cursor()
                    cursor.execute("SELECT data_de, data_ate FROM atestados_historico WHERE id_afastamento=? AND id!=?", (self.id_registro, id_atestado))
                    for ext in cursor.fetchall():
                        if ext[0] and ext[1]:
                            if dt_de <= datetime.strptime(ext[1], "%d/%m/%Y") and datetime.strptime(ext[0], "%d/%m/%Y") <= dt_ate:
                                conn.close()
                                return messagebox.showwarning("Aviso de Sobreposição", f"O período de {v_de} a {v_ate} sobrepõe-se a um existente ({ext[0]} a {ext[1]}).", parent=jan_ed)
                    conn.close()
                    d_v = (dt_ate - dt_de).days + 1
                except ValueError: return messagebox.showwarning("Erro", "Datas inválidas", parent=jan_ed)
                
                conn = conectar_db()
                conn.cursor().execute("UPDATE atestados_historico SET data_de=?, data_ate=?, qtd_dias=?, observacoes=? WHERE id=?", 
                                      (v_de, v_ate, d_v, ed_obs.get(), id_atestado))
                conn.commit()
                conn.close()
                carregar_lista_hist()
                self.sincronizar_totais_afastamento()
                jan_ed.destroy()
                
            tk.Button(jan_ed, text="💾 Guardar Alterações", command=salvar_ed_atestado, bg="#3B82F6", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", pady=5, padx=20).pack(pady=(0, 10))

        tree_hist.bind("<Double-1>", abrir_modal_editar_atestado)

        frame_add = ttk.LabelFrame(jan_hist, text=" Lançar Nova Prorrogação ", padding=15)
        frame_add.pack(fill="both", expand=True, padx=15, pady=5)
        tk.Label(frame_add, text="De:", bg="white").grid(row=0, column=0)
        ent_n_de = self.criar_calendario(frame_add, width=12)
        ent_n_de.grid(row=0, column=1, padx=5)
        tk.Label(frame_add, text="Até:", bg="white").grid(row=0, column=2, padx=(10,0))
        ent_n_ate = self.criar_calendario(frame_add, width=12)
        ent_n_ate.grid(row=0, column=3, padx=5)
        tk.Label(frame_add, text="Obs:", bg="white").grid(row=1, column=0, pady=10)
        ent_n_obs = tk.Entry(frame_add, width=35, relief="solid", bd=1)
        ent_n_obs.grid(row=1, column=1, columnspan=3, sticky="w", padx=5)

        data_ate_atual_str = self.ent_per_ate.get()
        if data_ate_atual_str:
            try:
                nova_data_de = datetime.strptime(data_ate_atual_str, "%d/%m/%Y") + timedelta(days=1)
                ent_n_de.set_date(nova_data_de)
                ent_n_ate.set_date(nova_data_de)
            except ValueError: pass

        def salvar_novo_atestado():
            de_v, ate_v, obs_v = ent_n_de.get(), ent_n_ate.get(), ent_n_obs.get()
            if not de_v or not ate_v: return messagebox.showwarning("Aviso", "Preencha as datas De e Até.", parent=jan_hist)
            try: 
                dt_de, dt_ate = datetime.strptime(de_v, "%d/%m/%Y"), datetime.strptime(ate_v, "%d/%m/%Y")
                if dt_ate < dt_de: return messagebox.showwarning("Erro", "A data 'Até' não pode ser anterior a 'De'.", parent=jan_hist)
                conn = conectar_db()
                cursor = conn.cursor()
                cursor.execute("SELECT data_de, data_ate FROM atestados_historico WHERE id_afastamento=?", (self.id_registro,))
                for ext in cursor.fetchall():
                    if ext[0] and ext[1]:
                        if dt_de <= datetime.strptime(ext[1], "%d/%m/%Y") and datetime.strptime(ext[0], "%d/%m/%Y") <= dt_ate:
                            conn.close()
                            return messagebox.showwarning("Aviso de Sobreposição", f"O período de {de_v} a {ate_v} sobrepõe-se a um existente ({ext[0]} a {ext[1]}).", parent=jan_hist)
                conn.close()
                dias_v = (dt_ate - dt_de).days + 1
            except ValueError: return messagebox.showwarning("Erro", "Datas inválidas", parent=jan_hist)

            conn = conectar_db()
            conn.cursor().execute("INSERT INTO atestados_historico (id_afastamento, data_de, data_ate, qtd_dias, observacoes) VALUES (?, ?, ?, ?, ?)", (self.id_registro, de_v, ate_v, dias_v, obs_v))
            conn.commit()
            conn.close()
            carregar_lista_hist()
            self.sincronizar_totais_afastamento() 
            ent_n_obs.delete(0, tk.END)
            try:
                nova_data_prox = datetime.strptime(ate_v, "%d/%m/%Y") + timedelta(days=1)
                ent_n_de.set_date(nova_data_prox)
                ent_n_ate.set_date(nova_data_prox)
            except ValueError: pass

        tk.Button(frame_add, text="Adicionar", command=salvar_novo_atestado, bg="#10B981", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2").grid(row=0, column=4, rowspan=2, padx=15, sticky="nsew")

    # ================= CARREGAR E SALVAR DADOS =================
    def buscar_dados(self):
        conn = conectar_db()
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM afastamentos WHERE id=?", (self.id_registro,))
        linha = cursor.fetchone()
        conn.close()
        return linha

    def preencher_data_segura(self, widget_calendario, valor_bd):
        widget_calendario.delete(0, tk.END)
        if valor_bd and str(valor_bd).strip() != "":
            try: widget_calendario.set_date(datetime.strptime(str(valor_bd), "%d/%m/%Y"))
            except Exception: widget_calendario.insert(0, str(valor_bd))

    def preencher_dados(self):
        ks = self.dados.keys()
        self.ent_empresa.insert(0, self.dados['empresa'] or "")
        self.ent_cod_empresa.insert(0, self.dados['cod_empresa'] if 'cod_empresa' in ks and self.dados['cod_empresa'] else "")
        self.ent_nome.insert(0, self.dados['nome_empregado'] or "")
        self.ent_cod_func.insert(0, self.dados['cod_empregado'] or "")
        self.combo_motivo.set(self.dados['tipo_afastamento'] or "Outros")
        
        for w, col in [(self.ent_per_de, 'periodo_de'), (self.ent_per_ate, 'periodo_ate'), (self.ent_dias, 'qtd_dias')]:
            w.config(state="normal")
            w.insert(0, self.dados[col] or "")
            w.config(state="readonly")

        conn = conectar_db()
        soma = conn.cursor().execute("SELECT sum(qtd_dias) FROM atestados_historico WHERE id_afastamento=?", (self.id_registro,)).fetchone()[0]
        conn.close()

        self.ent_soma_atestados.config(state="normal")
        self.ent_soma_atestados.insert(0, str(soma) if soma else "0")
        self.ent_soma_atestados.config(state="readonly")

        self.var_aborto.set((self.dados['aborto'] or 0) if 'aborto' in ks else 0)
        self.preencher_data_segura(self.ent_data_nasc, self.dados['data_nascimento'] if 'data_nascimento' in ks else "")
        self.combo_sal_fam.set(self.dados['salario_familia'] if 'salario_familia' in ks and self.dados['salario_familia'] else "")
        self.ent_obs_sal_fam.insert(0, self.dados['salario_familia_obs'] if 'salario_familia_obs' in ks and self.dados['salario_familia_obs'] else "")
        self.combo_imp_renda.set(self.dados['imposto_renda'] if 'imposto_renda' in ks and self.dados['imposto_renda'] else "")
        self.ent_obs_imp_renda.insert(0, self.dados['imposto_renda_obs'] if 'imposto_renda_obs' in ks and self.dados['imposto_renda_obs'] else "")

        self.txt_obs.insert("1.0", self.dados['observacoes'] if 'observacoes' in ks and self.dados['observacoes'] else "")
        self.var_obrig_concluida.set(self.dados['obrigacao_concluida'] or 0)
        self.var_retornou.set(self.dados['retornou'] or 0)
        self.combo_hiscre.set(self.dados['hiscre'] if 'hiscre' in ks and self.dados['hiscre'] else "")

        self.var_aposentado.set(1 if str(self.dados['aposentado']) in ("1", "Sim", "True") else 0)
        self.var_domestica.set(1 if 'domestica' in ks and str(self.dados['domestica']) in ("1", "Sim", "True") else 0)

        self.var_alerta_16.set((self.dados['alerta_16_tratado'] or 0) if 'alerta_16_tratado' in ks else 0)
        self.var_alerta_mat.set((self.dados['alerta_mat_tratado'] or 0) if 'alerta_mat_tratado' in ks else 0)
        self.var_alerta_retorno.set((self.dados['alerta_retorno_tratado'] or 0) if 'alerta_retorno_tratado' in ks else 0)
        self.var_alerta_ferias.set((self.dados['alerta_ferias_tratado'] or 0) if 'alerta_ferias_tratado' in ks else 0)
        
        self.toggle_aba_inss()
        self.toggle_maternidade()
        self.atualizar_cid_display()

        gov_auth = self.dados['gov_br_autorizacao'] if 'gov_br_autorizacao' in ks and self.dados['gov_br_autorizacao'] else ""
        cpf_val = self.dados['cpf'] if 'cpf' in ks and self.dados['cpf'] else ""
        gov_pwd = desencriptar_senha(self.dados['gov_br_senha'] if 'gov_br_senha' in ks and self.dados['gov_br_senha'] else "")

        self.ent_cpf.config(state="normal")
        self.ent_cpf.insert(0, cpf_val)
        self.combo_gov.set(gov_auth)
        if gov_auth == "Sim":
            self.ent_senha_gov.config(state="normal")
            self.ent_senha_gov.insert(0, gov_pwd)

        self.preencher_data_segura(self.ent_ult_dia, self.dados['ultimo_dia_trab'])
        self.preencher_data_segura(self.ent_obrigacao, self.dados['data_obrigacao'])
        self.preencher_data_segura(self.ent_16dia, self.dados['decimo_sexto_dia'])
        self.preencher_data_segura(self.ent_ret_data, self.dados['retorno_data'])
        self.preencher_data_segura(self.ent_estab, self.dados['estabilidade'])

        self.var_ctrl_ferias.set((self.dados['controlar_ferias'] or 0) if 'controlar_ferias' in ks else 0)
        self.preencher_data_segura(self.ent_aq_de, self.dados['per_aquisitivo_de'] if 'per_aquisitivo_de' in ks else "")
        self.preencher_data_segura(self.ent_novo_aq_de, self.dados['novo_per_aquisitivo_de'] if 'novo_per_aquisitivo_de' in ks else "")
        self.atualizar_fim_novo_aquisitivo()
        self.atualizar_termometro_ferias()

    def salvar_alteracoes(self):
        motivo = self.combo_motivo.get()
        if self.var_retornou.get() == 1 and motivo == "Acidente":
            if self.combo_hiscre.get() not in ["Sim", "Não"]:
                return messagebox.showwarning("Aviso", "Em caso de acidente é OBRIGATÓRIO informar se o HISCRE foi tirado.", parent=self.janela)

        old_nome, old_empresa = self.dados['nome_empregado'], self.dados['empresa']
        pwd_plana = self.ent_senha_gov.get()
        pwd_enc = encriptar_senha(pwd_plana) if pwd_plana else ""

        ret_str = self.ent_ret_data.get().strip()
        per_de_str = self.ent_per_de.get().strip()
        per_ate_str = self.ent_per_ate.get().strip()
        dias_str = self.ent_dias.get().strip()

        if motivo in ["Cárcere", "Militar"] and len(ret_str) == 10:
            try:
                ret_dt = datetime.strptime(ret_str, "%d/%m/%Y").date()
                fim_afast_dt = ret_dt - timedelta(days=1)
                
                if not per_de_str and self.ent_ult_dia.get().strip():
                    ult_dt = datetime.strptime(self.ent_ult_dia.get().strip(), "%d/%m/%Y").date()
                    per_de_str = (ult_dt + timedelta(days=1)).strftime("%d/%m/%Y")

                if per_de_str:
                    ini_dt = datetime.strptime(per_de_str, "%d/%m/%Y").date()
                    if fim_afast_dt >= ini_dt:
                        per_ate_str = fim_afast_dt.strftime("%d/%m/%Y")
                        dias_str = str((fim_afast_dt - ini_dt).days + 1)
            except ValueError:
                pass

        dias_val = int(dias_str) if dias_str.isdigit() else 0
        if dias_val > 15:
            self.var_ctrl_ferias.set(1)

        ctrl_ferias = self.var_ctrl_ferias.get()
        aq_de = self.ent_aq_de.get().strip() if ctrl_ferias == 1 else ""
        aq_ate = self.ent_aq_ate.get().strip() if ctrl_ferias == 1 else ""
        is_aposentado = str(self.var_aposentado.get()) in ("1", "Sim", "True")
        
        tela_novo_de = self.ent_novo_aq_de.get().strip()
        tela_novo_ate = self.ent_novo_aq_ate.get().strip()
        
        novo_aq_de = ""
        novo_aq_ate = ""

        if ctrl_ferias == 1 and len(tela_novo_de) == 10:
            if self.dias_beneficio_aquisitivo >= 180 or motivo in ["Cárcere", "Militar"]:
                novo_aq_de = tela_novo_de
                novo_aq_ate = tela_novo_ate
                
            elif is_aposentado and dias_val > 30 and motivo in ["Doença", "Acidente", "Outros"]:
                considerar_novo = messagebox.askyesno(
                    "Confirmação de Período Aquisitivo",
                    f"O cálculo indica que o novo período aquisitivo iniciará em {tela_novo_de}.\n\n"
                    f"Deseja gravar este novo período aquisitivo na base de dados?\n"
                    f"(Clique em 'Não' para manter o período aquisitivo original)",
                    parent=self.janela
                )
                if considerar_novo:
                    novo_aq_de = tela_novo_de
                    novo_aq_ate = tela_novo_ate
                else:
                    self.ent_novo_aq_de.config(state="normal")
                    self.ent_novo_aq_de.delete(0, tk.END)
                    self.ent_novo_aq_ate.config(state="normal")
                    self.ent_novo_aq_ate.delete(0, tk.END)
                    self.ent_novo_aq_ate.config(state="readonly")
            else:
                 novo_aq_de = tela_novo_de
                 novo_aq_ate = tela_novo_ate

        if ctrl_ferias == 1 and (not aq_de or len(aq_de) != 10):
            messagebox.showwarning("Aviso", "Como o afastamento é superior a 15 dias, é obrigatório preencher o 'Início Per. Aquisitivo' para o controle de Férias.", parent=self.janela)
            self.notebook.select(self.aba_resumo)
            return

        try:
            conn = conectar_db()
            cursor = conn.cursor()

            if motivo == "Cárcere" and len(ret_str) == 10 and per_de_str and per_ate_str and dias_val > 0:
                cursor.execute("SELECT id FROM atestados_historico WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
                hist_rows = cursor.fetchall()
                if not hist_rows:
                    cursor.execute(
                        "INSERT INTO atestados_historico (id_afastamento, data_de, data_ate, qtd_dias, observacoes) VALUES (?, ?, ?, ?, ?)",
                        (self.id_registro, per_de_str, per_ate_str, dias_val, "Período de reclusão (fechado no retorno)")
                    )
                elif len(hist_rows) == 1:
                    cursor.execute(
                        "UPDATE atestados_historico SET data_de=?, data_ate=?, qtd_dias=? WHERE id=?",
                        (per_de_str, per_ate_str, dias_val, hist_rows[0][0])
                    )

            cursor.execute('''
                UPDATE afastamentos SET
                    empresa=?, cod_empresa=?, cod_empregado=?, nome_empregado=?, tipo_afastamento=?, 
                    ultimo_dia_trab=?, periodo_de=?, periodo_ate=?, qtd_dias=?,
                    data_obrigacao=?, obrigacao_concluida=?,
                    decimo_sexto_dia=?, retornou=?, retorno_data=?, estabilidade=?, 
                    observacoes=?, hiscre=?, gov_br_autorizacao=?, gov_br_senha=?, cpf=?, 
                    aposentado=?, domestica=?, alerta_16_tratado=?, alerta_mat_tratado=?, alerta_retorno_tratado=?,
                    aborto=?, data_nascimento=?, salario_familia=?, salario_familia_obs=?, imposto_renda=?, imposto_renda_obs=?,
                    controlar_ferias=?, per_aquisitivo_de=?, per_aquisitivo_ate=?,
                    novo_per_aquisitivo_de=?, novo_per_aquisitivo_ate=?, alerta_ferias_tratado=?
                WHERE id=?
            ''', (
                self.ent_empresa.get().strip(), self.ent_cod_empresa.get().strip(), self.ent_cod_func.get().strip(), self.ent_nome.get().strip(), motivo,
                self.ent_ult_dia.get(), per_de_str, per_ate_str, dias_str,
                self.ent_obrigacao.get(), self.var_obrig_concluida.get(),
                self.ent_16dia.get(), self.var_retornou.get(), ret_str, self.ent_estab.get(), 
                self.txt_obs.get("1.0", tk.END).strip(), self.combo_hiscre.get(), self.combo_gov.get(), 
                pwd_enc, self.ent_cpf.get(), self.var_aposentado.get(), self.var_domestica.get(), 
                self.var_alerta_16.get(), self.var_alerta_mat.get(), self.var_alerta_retorno.get(),
                self.var_aborto.get(), 
                self.ent_data_nasc.get() if motivo == "Maternidade" else "",
                self.combo_sal_fam.get() if motivo == "Maternidade" else "",
                self.ent_obs_sal_fam.get().strip() if motivo == "Maternidade" else "",
                self.combo_imp_renda.get() if motivo == "Maternidade" else "",
                self.ent_obs_imp_renda.get().strip() if motivo == "Maternidade" else "",
                ctrl_ferias, aq_de, aq_ate, novo_aq_de, novo_aq_ate, self.var_alerta_ferias.get(),
                self.id_registro
            ))

            if motivo == "Maternidade":
                cursor.execute('''
                    UPDATE gestantes SET empresa=?, cod_empresa=?, cod_empregado=?, nome_empregado=?
                    WHERE nome_empregado=? AND empresa=?
                ''', (self.ent_empresa.get().strip(), self.ent_cod_empresa.get().strip(), self.ent_cod_func.get().strip(), self.ent_nome.get().strip(), old_nome, old_empresa))

            conn.commit()
            conn.close()

            if ctrl_ferias == 1 and aq_de and self.var_retornou.get() == 0:
                sincronizar_aquisitivo_funcionario(
                    self.ent_cpf.get().strip(), self.ent_cod_empresa.get().strip(),
                    self.ent_cod_func.get().strip(), ctrl_ferias, aq_de, aq_ate
                )
            
            messagebox.showinfo("Sucesso", "Registo atualizado com sucesso!", parent=self.janela)
            self.callback_atualizar()
            self.janela.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar: {str(e)}", parent=self.janela)

    def atualizar_fim_novo_aquisitivo(self, event=None):
        dt_de = self.ent_novo_aq_de.get().strip()
        dt_ate = calcular_fim_aquisitivo(dt_de)
        self.ent_novo_aq_ate.config(state="normal")
        self.ent_novo_aq_ate.delete(0, tk.END)
        if dt_ate:
            self.ent_novo_aq_ate.insert(0, dt_ate)
        self.ent_novo_aq_ate.config(state="readonly")

    def sincronizar_totais_afastamento(self):
        conn = conectar_db()
        cursor = conn.cursor()
        cursor.execute("SELECT data_de, data_ate FROM atestados_historico WHERE id_afastamento=? ORDER BY id ASC", (self.id_registro,))
        atestados = cursor.fetchall()
        
        motivo_atual = self.combo_motivo.get()

        if atestados:
            datas_de, datas_ate = [], []
            for a in atestados:
                try: datas_de.append(datetime.strptime(a[0], "%d/%m/%Y"))
                except: pass
                try: datas_ate.append(datetime.strptime(a[1], "%d/%m/%Y"))
                except: pass
            
            if datas_de and datas_ate:
                min_de, max_ate = min(datas_de), max(datas_ate)
                total_dias = (max_ate - min_de).days + 1
                str_de, str_ate = min_de.strftime("%d/%m/%Y"), max_ate.strftime("%d/%m/%Y")
                
                cursor.execute("UPDATE afastamentos SET periodo_de=?, periodo_ate=?, qtd_dias=? WHERE id=?", (str_de, str_ate, total_dias, self.id_registro))
                
                for widget, val in [(self.ent_per_de, str_de), (self.ent_per_ate, str_ate), (self.ent_dias, str(total_dias))]:
                    widget.config(state="normal")
                    widget.delete(0, tk.END)
                    widget.insert(0, val)
                    widget.config(state="readonly")
        else:
            if motivo_atual in ["Cárcere", "Militar"]:
                ult_trab = self.ent_ult_dia.get().strip()
                if ult_trab:
                    try:
                        dt_inicio_afast = datetime.strptime(ult_trab, "%d/%m/%Y") + timedelta(days=1)
                        str_de = dt_inicio_afast.strftime("%d/%m/%Y")
                        cursor.execute("UPDATE afastamentos SET periodo_de=?, periodo_ate='', qtd_dias=0 WHERE id=?", (str_de, self.id_registro))
                        
                        self.ent_per_de.config(state="normal")
                        self.ent_per_de.delete(0, tk.END)
                        self.ent_per_de.insert(0, str_de)
                        self.ent_per_de.config(state="readonly")
                    except: pass
            else:
                cursor.execute("UPDATE afastamentos SET periodo_de='', periodo_ate='', qtd_dias='' WHERE id=?", (self.id_registro,))
                for obj in [self.ent_per_de, self.ent_per_ate, self.ent_dias]:
                    obj.config(state="normal")
                    obj.delete(0, tk.END)
                    obj.config(state="readonly")

        cursor.execute("SELECT sum(qtd_dias) FROM atestados_historico WHERE id_afastamento=?", (self.id_registro,))
        soma = cursor.fetchone()[0]
        self.ent_soma_atestados.config(state="normal")
        self.ent_soma_atestados.delete(0, tk.END)
        self.ent_soma_atestados.insert(0, str(soma) if soma else "0")
        self.ent_soma_atestados.config(state="readonly")
                
        conn.commit()
        conn.close()
        self.callback_atualizar()

    def sincronizar_com_beneficio(self, tipo):
        tree = self.tree_an if tipo == 'analise' else self.tree_per
        selec = tree.selection()
        if not selec:
            return messagebox.showwarning("Aviso", "Selecione uma linha na tabela acima para poder sincronizar.", parent=self.janela)

        valores = tree.item(selec[0], "values")
        num_beneficio = valores[3] if tipo == 'analise' else valores[5]
        situacao = valores[4] if tipo == 'analise' else valores[6]
        concessao_ate = valores[6] if tipo == 'analise' else valores[8]

        if situacao not in ["Concedido", "Concedida"]:
            return messagebox.showwarning("Aviso", "Apenas benefícios com situação 'Concedido' podem gerar extensão de atestado.", parent=self.janela)
        if not concessao_ate or concessao_ate == "-":
            return messagebox.showwarning("Aviso", "O benefício selecionado não possui Data Final (Até).", parent=self.janela)

        data_atual_str = self.ent_per_ate.get().strip()
        if not data_atual_str:
            return messagebox.showwarning("Aviso", "O afastamento não possui uma data final atual para usar como base.", parent=self.janela)

        try:
            dt_atual_fim = datetime.strptime(data_atual_str, "%d/%m/%Y")
            dt_nova_fim = datetime.strptime(concessao_ate, "%d/%m/%Y")
            if dt_nova_fim <= dt_atual_fim:
                return messagebox.showinfo("Aviso", f"O benefício ({concessao_ate}) não é posterior ao término atual ({data_atual_str}).", parent=self.janela)

            dt_novo_inicio = dt_atual_fim + timedelta(days=1)
            dias_extensao = (dt_nova_fim - dt_novo_inicio).days + 1
            if dias_extensao > 0:
                str_novo_inicio = dt_novo_inicio.strftime("%d/%m/%Y")
                obs = f"Ajuste de data segundo beneficio nº {num_beneficio}" if num_beneficio and num_beneficio != "-" else "Ajuste de data segundo beneficio INSS"
                if messagebox.askyesno("Confirmar Sincronização", f"Gerar atestado de ajuste do dia {str_novo_inicio} até {concessao_ate} ({dias_extensao} dias)?", parent=self.janela):
                    conn = conectar_db()
                    conn.cursor().execute("INSERT INTO atestados_historico (id_afastamento, data_de, data_ate, qtd_dias, observacoes) VALUES (?, ?, ?, ?, ?)", 
                                          (self.id_registro, str_novo_inicio, concessao_ate, dias_extensao, obs))
                    conn.commit()
                    conn.close()
                    self.sincronizar_totais_afastamento()
                    messagebox.showinfo("Sucesso", "Afastamento estendido com sucesso!", parent=self.janela)
        except ValueError:
            messagebox.showerror("Erro", "Erro ao processar as datas.", parent=self.janela)