import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import json
import os
import csv
import platform
import subprocess
from database import (
    conectar_db, calcular_fim_aquisitivo, buscar_dados_funcionario_existente,
    calcular_dias_beneficio_aquisitivo, sincronizar_aquisitivo_funcionario
)

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT

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

class TelaCadastro:
    def __init__(self, parent, callback_atualizar):
        self.callback_atualizar = callback_atualizar
        self.cids_lista = []
        self.map_gestantes = {}
        
        self.janela_cad = tk.Toplevel(parent)
        self.janela_cad.title("Registar Novo Afastamento")
        self.janela_cad.geometry("880x830") 
        self.janela_cad.grab_set() 
        self.janela_cad.resizable(False, False)
        self.janela_cad.configure(bg="#F3F4F6")

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TLabelframe", background="#FFFFFF", borderwidth=2)
        style.configure("TLabelframe.Label", background="#FFFFFF", font=("Segoe UI", 10, "bold"), foreground="#2563EB")
        style.configure("TSeparator", background="#E5E7EB")

        fonte_label = ("Segoe UI", 9, "bold")
        fonte_entry = ("Segoe UI", 9)

        # ================= 1. TIPO DE AFASTAMENTO E DADOS BÁSICOS =================
        box_dados = ttk.LabelFrame(self.janela_cad, text=" 👤 Tipo de Afastamento e Dados Básicos ", padding=15)
        box_dados.pack(fill="x", padx=20, pady=(12, 6))

        tk.Label(box_dados, text="Motivo:", bg="white", font=fonte_label, fg="#DC2626").grid(row=0, column=0, sticky="w", pady=4)
        self.combo_motivo = ttk.Combobox(box_dados, values=["Doença", "Acidente", "Cárcere", "Militar", "Maternidade", "Outros"], width=31, state="readonly", font=fonte_entry)
        self.combo_motivo.grid(row=0, column=1, padx=(5, 15), sticky="w", pady=4)
        self.combo_motivo.set("Doença")

        self.lbl_vinc_gestante = tk.Label(box_dados, text="Vincular Gestante:", bg="white", font=fonte_label, fg="#DB2777")
        self.combo_gestante = ttk.Combobox(box_dados, width=45, state="readonly", font=fonte_entry)

        def carregar_gestantes():
            self.map_gestantes.clear()
            opcoes = [""]
            conn = conectar_db()
            c = conn.cursor()
            c.execute("SELECT id, nome_empregado, empresa, cod_empresa, cod_empregado, estabilidade FROM gestantes WHERE status='Gestante'")
            for r in c.fetchall():
                display_text = f"{r[1]} - {r[2]}"
                self.map_gestantes[display_text] = {
                    'id': r[0], 'nome_real': r[1], 'empresa': r[2], 
                    'cod_empresa': r[3], 'cod_func': r[4], 'estab': r[5]
                }
                opcoes.append(display_text)
            conn.close()
            self.combo_gestante['values'] = opcoes

        def auto_preencher_gestante(e=None):
            selecao = self.combo_gestante.get()
            if selecao in self.map_gestantes:
                dados_g = self.map_gestantes[selecao]
                self.ent_empresa.delete(0, tk.END)
                self.ent_empresa.insert(0, dados_g['empresa'] or "")
                self.ent_cod_empresa.delete(0, tk.END)
                self.ent_cod_empresa.insert(0, dados_g['cod_empresa'] or "")
                self.ent_nome.delete(0, tk.END)
                self.ent_nome.insert(0, dados_g['nome_real']) 
                self.ent_cod_func.delete(0, tk.END)
                self.ent_cod_func.insert(0, dados_g['cod_func'] or "")
                self.tentar_vincular_historico_funcionario()

        self.combo_gestante.bind("<<ComboboxSelected>>", auto_preencher_gestante)

        ttk.Separator(box_dados, orient='horizontal').grid(row=1, column=0, columnspan=4, sticky="ew", pady=8)

        tk.Label(box_dados, text="CPF do Func.:", bg="white", font=fonte_label).grid(row=2, column=0, sticky="w", pady=4)
        f_cpf = tk.Frame(box_dados, bg="white")
        f_cpf.grid(row=2, column=1, columnspan=3, sticky="w", padx=(5, 15), pady=4)
        
        self.ent_cpf = tk.Entry(f_cpf, width=16, font=fonte_entry, relief="solid", bd=1)
        self.ent_cpf.pack(side="left")
        self.ent_cpf.bind("<KeyRelease>", self.formatar_cpf)
        self.ent_cpf.bind("<FocusOut>", lambda e: self.tentar_vincular_historico_funcionario())

        btn_termo = tk.Button(f_cpf, text="📄 Emitir Autorização GOV.BR", command=self.gerar_termo_govbr, bg="#3B82F6", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2")
        btn_termo.pack(side="left", padx=10)

        tk.Label(box_dados, text="Empresa:", bg="white", font=fonte_label).grid(row=3, column=0, sticky="w", pady=4)
        self.ent_empresa = tk.Entry(box_dados, width=34, font=fonte_entry, relief="solid", bd=1)
        self.ent_empresa.grid(row=3, column=1, padx=(5, 15), sticky="w", pady=4)

        tk.Label(box_dados, text="Cód. Empresa:", bg="white", font=fonte_label).grid(row=3, column=2, sticky="e", padx=(10, 5), pady=4)
        self.ent_cod_empresa = tk.Entry(box_dados, width=12, font=fonte_entry, relief="solid", bd=1)
        self.ent_cod_empresa.grid(row=3, column=3, sticky="w", pady=4)
        self.ent_cod_empresa.bind("<FocusOut>", lambda e: self.tentar_vincular_historico_funcionario())

        tk.Label(box_dados, text="Funcionário:", bg="white", font=fonte_label).grid(row=4, column=0, sticky="w", pady=4)
        self.ent_nome = tk.Entry(box_dados, width=34, font=fonte_entry, relief="solid", bd=1)
        self.ent_nome.grid(row=4, column=1, padx=(5, 15), sticky="w", pady=4)

        tk.Label(box_dados, text="Cód. Func.:", bg="white", font=fonte_label).grid(row=4, column=2, sticky="e", padx=(10, 5), pady=4)
        self.ent_cod_func = tk.Entry(box_dados, width=12, font=fonte_entry, relief="solid", bd=1)
        self.ent_cod_func.grid(row=4, column=3, sticky="w", pady=4)
        self.ent_cod_func.bind("<FocusOut>", lambda e: self.tentar_vincular_historico_funcionario())

        f_chk = tk.Frame(box_dados, bg="white")
        f_chk.grid(row=5, column=0, columnspan=2, sticky="w", pady=4)
        self.var_aposentado = tk.IntVar()
        tk.Checkbutton(f_chk, text="Aposentado(a)", variable=self.var_aposentado, bg="white", font=fonte_label, fg="#EF4444").pack(side="left", padx=(0, 10))
        self.var_domestica = tk.IntVar()
        tk.Checkbutton(f_chk, text="Doméstica", variable=self.var_domestica, bg="white", font=fonte_label, fg="#8B5CF6").pack(side="left")

        tk.Label(box_dados, text="Últ. Dia Trab:", bg="white", font=fonte_label).grid(row=6, column=0, sticky="w", pady=4)
        self.ent_ult_dia = self.criar_calendario(box_dados)
        self.ent_ult_dia.grid(row=6, column=1, padx=5, sticky="w", pady=4)
        self.ent_ult_dia.delete(0, tk.END)

        tk.Label(box_dados, text="Data Obrigação:", bg="white", font=fonte_label).grid(row=6, column=2, sticky="e", padx=(10, 5), pady=4)
        self.ent_obrigacao = self.criar_calendario(box_dados)
        self.ent_obrigacao.grid(row=6, column=3, sticky="w", pady=4)
        self.ent_obrigacao.delete(0, tk.END)

        # --- BLOCO MATERNIDADE ---
        self.frame_mat = tk.Frame(box_dados, bg="#FFF1F2", bd=1, relief="solid", padx=10, pady=6)
        self.var_aborto = tk.IntVar()
        tk.Checkbutton(self.frame_mat, text="Afastamento por Aborto", variable=self.var_aborto, bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=0, column=0, sticky="w", pady=3)
        tk.Label(self.frame_mat, text="Data Nascimento:", bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=0, column=1, sticky="e", padx=(10, 5), pady=3)
        self.ent_data_nasc = self.criar_calendario(self.frame_mat, width=12)
        self.ent_data_nasc.grid(row=0, column=2, sticky="w", pady=3)
        self.ent_data_nasc.delete(0, tk.END)

        tk.Label(self.frame_mat, text="Salário Família?:", bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=1, column=0, sticky="w", pady=3)
        self.combo_sal_fam = ttk.Combobox(self.frame_mat, values=["", "Sim", "Não"], width=10, state="readonly", font=fonte_entry)
        self.combo_sal_fam.grid(row=1, column=1, sticky="w", padx=5, pady=3)
        tk.Label(self.frame_mat, text="Obs Sal. Família:", bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=1, column=2, sticky="e", padx=(10, 5), pady=3)
        self.ent_obs_sal_fam = tk.Entry(self.frame_mat, width=28, font=fonte_entry, relief="solid", bd=1)
        self.ent_obs_sal_fam.grid(row=1, column=3, sticky="w", padx=5, pady=3)

        tk.Label(self.frame_mat, text="Imposto Renda?:", bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=2, column=0, sticky="w", pady=3)
        self.combo_imp_renda = ttk.Combobox(self.frame_mat, values=["", "Sim", "Não"], width=10, state="readonly", font=fonte_entry)
        self.combo_imp_renda.grid(row=2, column=1, sticky="w", padx=5, pady=3)
        tk.Label(self.frame_mat, text="Obs Imp. Renda:", bg="#FFF1F2", font=fonte_label, fg="#E11D48").grid(row=2, column=2, sticky="e", padx=(10, 5), pady=3)
        self.ent_obs_imp_renda = tk.Entry(self.frame_mat, width=28, font=fonte_entry, relief="solid", bd=1)
        self.ent_obs_imp_renda.grid(row=2, column=3, sticky="w", padx=5, pady=3)

        def toggle_maternidade(e=None):
            if self.combo_motivo.get() == "Maternidade":
                self.lbl_vinc_gestante.grid(row=0, column=2, sticky="e", padx=(10, 5), pady=4)
                self.combo_gestante.grid(row=0, column=3, sticky="w", pady=4)
                self.frame_mat.grid(row=7, column=0, columnspan=4, sticky="ew", pady=(8, 2))
                self.janela_cad.geometry("880x940")
                carregar_gestantes()
            else:
                self.lbl_vinc_gestante.grid_forget()
                self.combo_gestante.grid_forget()
                self.frame_mat.grid_forget()
                self.janela_cad.geometry("880x830")

        self.combo_motivo.bind("<<ComboboxSelected>>", toggle_maternidade)

        # ================= 2. PERÍODO DO AFASTAMENTO E CONTROLE DE FÉRIAS =================
        box_periodo = ttk.LabelFrame(self.janela_cad, text=" 📅 Período do Afastamento e Controle de Férias ", padding=15)
        box_periodo.pack(fill="x", padx=20, pady=6)

        tk.Label(box_periodo, text="Período De:", bg="white", font=fonte_label).grid(row=0, column=0, sticky="w", pady=4)
        self.ent_per_de = self.criar_calendario(box_periodo, cor_fundo='#10B981')
        self.ent_per_de.grid(row=0, column=1, padx=5, sticky="w", pady=4)
        self.ent_per_de.delete(0, tk.END)
        self.ent_per_de.bind("<<DateEntrySelected>>", lambda e: self.calcular_por_datas("de"))
        self.ent_per_de.bind("<KeyRelease>", lambda e: self.calcular_por_datas("de"))

        tk.Label(box_periodo, text="Período Até:", bg="white", font=fonte_label).grid(row=0, column=2, sticky="e", padx=(15, 5), pady=4)
        self.ent_per_ate = self.criar_calendario(box_periodo, cor_fundo='#10B981')
        self.ent_per_ate.grid(row=0, column=3, sticky="w", pady=4)
        self.ent_per_ate.delete(0, tk.END)
        self.ent_per_ate.bind("<<DateEntrySelected>>", lambda e: self.calcular_por_datas("ate"))
        self.ent_per_ate.bind("<KeyRelease>", lambda e: self.calcular_por_datas("ate"))

        tk.Label(box_periodo, text="Qtd Dias (Auto):", bg="white", font=fonte_label).grid(row=1, column=0, sticky="w", pady=4)
        self.ent_dias = tk.Entry(box_periodo, width=15, font=("Segoe UI", 10, "bold"), justify="center", bg="#FFFFFF", relief="solid", bd=1)
        self.ent_dias.grid(row=1, column=1, padx=5, sticky="w", pady=4)
        self.ent_dias.bind("<KeyRelease>", self.calcular_por_dias)

        tk.Label(box_periodo, text="16º Dia (Auto):", bg="white", font=fonte_label).grid(row=1, column=2, sticky="e", padx=(15, 5), pady=4)
        self.ent_16dia = self.criar_calendario(box_periodo, cor_fundo='#F59E0B')
        self.ent_16dia.grid(row=1, column=3, sticky="w", pady=4)
        self.ent_16dia.delete(0, tk.END)

        # --- SUB-BLOCO: CONTROLE DE PERÍODO AQUISITIVO (180 DIAS BENEFÍCIO) ---
        ttk.Separator(box_periodo, orient='horizontal').grid(row=2, column=0, columnspan=5, sticky="ew", pady=(10, 8))

        self.var_ctrl_ferias = tk.IntVar(value=0)
        chk_ferias = tk.Checkbutton(
            box_periodo, text="🏖️ Controlar Período Aquisitivo de Férias (Limite 180 dias Benefício Concedido)",
            variable=self.var_ctrl_ferias, command=self.toggle_campos_ferias,
            bg="white", font=fonte_label, fg="#0284C7"
        )
        chk_ferias.grid(row=3, column=0, columnspan=4, sticky="w", pady=(0, 6))

        self.f_aquisitivo = tk.Frame(box_periodo, bg="#F0F9FF", bd=1, relief="solid", padx=10, pady=6)
        self.f_aquisitivo.grid(row=4, column=0, columnspan=5, sticky="ew")

        tk.Label(self.f_aquisitivo, text="Início Per. Aquisitivo:", bg="#F0F9FF", font=fonte_label, fg="#0369A1").pack(side="left")
        self.ent_aq_de = self.criar_calendario(self.f_aquisitivo, width=12, cor_fundo='#0284C7')
        self.ent_aq_de.pack(side="left", padx=(5, 15))
        self.ent_aq_de.delete(0, tk.END)
        self.ent_aq_de.bind("<<DateEntrySelected>>", self.atualizar_fim_aquisitivo)
        self.ent_aq_de.bind("<KeyRelease>", self.atualizar_fim_aquisitivo)

        tk.Label(self.f_aquisitivo, text="Fim Per. Aquisitivo (Auto):", bg="#F0F9FF", font=fonte_label, fg="#0369A1").pack(side="left")
        self.ent_aq_ate = tk.Entry(self.f_aquisitivo, width=12, font=fonte_entry, relief="solid", bd=1, state="readonly", justify="center")
        self.ent_aq_ate.pack(side="left", padx=(5, 15))

        self.lbl_resumo_aq = tk.Label(self.f_aquisitivo, text="Benefícios Concedidos no ciclo: 0 / 180 dias", bg="#F0F9FF", font=("Segoe UI", 9, "bold"), fg="#0369A1")
        self.lbl_resumo_aq.pack(side="left", padx=5)

        self.toggle_campos_ferias()

        # ================= 3. DIAGNÓSTICOS MÉDICOS (BUSCA ACIMA DO GRID) =================
        box_cid = ttk.LabelFrame(self.janela_cad, text=" 🏥 Diagnósticos Médicos (CIDs no Atestado) ", padding=12)
        box_cid.pack(fill="x", padx=20, pady=6)

        f_ins_cid = tk.Frame(box_cid, bg="white")
        f_ins_cid.pack(fill="x", pady=(0, 10))

        tk.Label(f_ins_cid, text="Código:", bg="white", font=fonte_label).pack(side="left")
        self.ent_cid_cod = tk.Entry(f_ins_cid, width=9, font=fonte_entry, relief="solid", bd=1)
        self.ent_cid_cod.pack(side="left", padx=4)
        self.ent_cid_cod.bind("<Return>", lambda e: self.acionar_busca_cid())

        btn_busca = tk.Button(f_ins_cid, text="🔍 Buscar", command=self.acionar_busca_cid, bg="#3B82F6", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2")
        btn_busca.pack(side="left", padx=(0, 10))

        tk.Label(f_ins_cid, text="Descrição:", bg="white", font=fonte_label).pack(side="left")
        self.ent_cid_desc = tk.Entry(f_ins_cid, width=44, font=fonte_entry, relief="solid", bd=1)
        self.ent_cid_desc.pack(side="left", padx=4)
        self.ent_cid_desc.bind("<Return>", lambda e: self.adicionar_cid_lista())

        btn_add_cid = tk.Button(f_ins_cid, text="➕ Adicionar CID", command=self.adicionar_cid_lista, bg="#10B981", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2")
        btn_add_cid.pack(side="left", padx=8)

        self.tree_cid = ttk.Treeview(box_cid, columns=("Código", "Descrição"), show="headings", height=4)
        self.tree_cid.heading("Código", text="Cód. CID")
        self.tree_cid.heading("Descrição", text="Descrição do Diagnóstico")
        self.tree_cid.column("Código", width=90, anchor="center")
        self.tree_cid.column("Descrição", width=700, anchor="w")
        self.tree_cid.pack(fill="x")

        self.tree_cid.bind("<Shift-Delete>", self.excluir_cid_selecionado)
        self.tree_cid.bind("<Double-1>", self.abrir_modal_editar_cid)

        # ================= 4. BOTÕES (LIMPAR / SALVAR) =================
        f_botoes = tk.Frame(self.janela_cad, bg="#F3F4F6")
        f_botoes.pack(fill="x", padx=20, pady=(6, 12))

        btn_limpar = tk.Button(f_botoes, text="🧹 Limpar Tela", command=self.limpar_tela, 
                               bg="#9CA3AF", fg="white", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", pady=8, width=15)
        btn_limpar.pack(side="left", padx=(0, 10))

        btn_salvar = tk.Button(f_botoes, text="✔ Guardar Novo Registo", command=self.salvar_registro, 
                               bg="#10B981", fg="white", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", pady=8)
        btn_salvar.pack(side="left", fill="x", expand=True)

    def verificar_obrigatoriedade_ferias(self, dias):
        if dias > 15:
            if self.var_ctrl_ferias.get() == 0:
                self.var_ctrl_ferias.set(1)
                self.toggle_campos_ferias()
        else:
            if self.var_ctrl_ferias.get() == 1:
                self.var_ctrl_ferias.set(0)
                self.toggle_campos_ferias()

    def limpar_tela(self):
        if not messagebox.askyesno("Limpar Tela", "Tem certeza que deseja limpar todos os campos?", parent=self.janela_cad):
            return
            
        self.combo_motivo.set("Doença")
        self.combo_gestante.set("")
        self.ent_empresa.delete(0, tk.END)
        self.ent_cod_empresa.delete(0, tk.END)
        self.ent_nome.delete(0, tk.END)
        self.ent_cod_func.delete(0, tk.END)
        self.ent_cpf.delete(0, tk.END)
        
        self.var_aposentado.set(0)
        self.var_domestica.set(0)
        
        self.ent_ult_dia.delete(0, tk.END)
        self.ent_obrigacao.delete(0, tk.END)
        
        self.var_aborto.set(0)
        self.ent_data_nasc.delete(0, tk.END)
        self.combo_sal_fam.set("")
        self.ent_obs_sal_fam.delete(0, tk.END)
        self.combo_imp_renda.set("")
        self.ent_obs_imp_renda.delete(0, tk.END)
        
        self.toggle_maternidade() 
        
        self.ent_per_de.delete(0, tk.END)
        self.ent_per_ate.delete(0, tk.END)
        self.ent_dias.delete(0, tk.END)
        self.ent_16dia.delete(0, tk.END)
        
        self.var_ctrl_ferias.set(0)
        self.toggle_campos_ferias()
        self.ent_aq_de.delete(0, tk.END)
        self.ent_aq_ate.config(state="normal")
        self.ent_aq_ate.delete(0, tk.END)
        self.ent_aq_ate.config(state="readonly")
        self.lbl_resumo_aq.config(text="Benefícios Concedidos no ciclo: 0 / 180 dias", fg="#0369A1")
        
        self.ent_cid_cod.delete(0, tk.END)
        self.ent_cid_desc.delete(0, tk.END)
        self.cids_lista.clear()
        self.atualizar_tabela_cids()
        
        self.ent_cpf.focus_set()

    def toggle_campos_ferias(self):
        if self.var_ctrl_ferias.get() == 1:
            self.f_aquisitivo.grid()
        else:
            self.f_aquisitivo.grid_remove()

    def atualizar_fim_aquisitivo(self, event=None):
        dt_de = self.ent_aq_de.get().strip()
        dt_ate = calcular_fim_aquisitivo(dt_de)
        self.ent_aq_ate.config(state="normal")
        self.ent_aq_ate.delete(0, tk.END)
        if dt_ate:
            self.ent_aq_ate.insert(0, dt_ate)
        self.ent_aq_ate.config(state="readonly")
        self.atualizar_resumo_beneficios_anteriores()

    def atualizar_resumo_beneficios_anteriores(self):
        aq_de = self.ent_aq_de.get().strip()
        aq_ate = self.ent_aq_ate.get().strip()
        if len(aq_de) == 10 and len(aq_ate) == 10:
            dias, qtd_af = calcular_dias_beneficio_aquisitivo(
                self.ent_cpf.get(), self.ent_cod_empresa.get(), self.ent_cod_func.get(), aq_de, aq_ate
            )
            cor = "#DC2626" if dias >= 180 else ("#D97706" if dias >= 150 else "#0369A1")
            self.lbl_resumo_aq.config(text=f"Benefícios Concedidos no ciclo: {dias} / 180 dias ({qtd_af} reg.)", fg=cor)
        else:
            self.lbl_resumo_aq.config(text="Benefícios Concedidos no ciclo: 0 / 180 dias", fg="#0369A1")

    def tentar_vincular_historico_funcionario(self):
        cpf_val = self.ent_cpf.get().strip()
        if not cpf_val or len(cpf_val) < 11:
            return

        conn = conectar_db()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT empresa, cod_empresa, nome_empregado, cod_empregado, 
                   controlar_ferias, per_aquisitivo_de, per_aquisitivo_ate,
                   novo_per_aquisitivo_de, novo_per_aquisitivo_ate, aposentado, domestica
            FROM afastamentos 
            WHERE cpf=? 
            ORDER BY id DESC LIMIT 1
        """, (cpf_val,))
        
        reg = cursor.fetchone()
        conn.close()
        
        print("--- DEBUG HISTÓRICO CPF ---")
        print("Registro encontrado:", reg)

        if reg:
            empresa, cod_emp, nome, cod_func, ctrl_f, p_de, p_ate, n_de, n_ate, apos, dom = reg
            print(f"Período Normal: {p_de} a {p_ate}")
            print(f"Novo Período (Retorno): {n_de} a {n_ate}")
            
            if n_de and len(str(n_de).strip()) == 10:
                periodo_ini_usar = n_de
            else:
                periodo_ini_usar = p_de

            if not self.ent_nome.get().strip() and nome:
                self.ent_nome.insert(0, nome)
            if not self.ent_empresa.get().strip() and empresa:
                self.ent_empresa.insert(0, empresa)
            if not self.ent_cod_empresa.get().strip() and cod_emp:
                self.ent_cod_empresa.insert(0, cod_emp)
            if not self.ent_cod_func.get().strip() and cod_func:
                self.ent_cod_func.insert(0, cod_func)

            if periodo_ini_usar:
                self.var_ctrl_ferias.set(ctrl_f if ctrl_f is not None else 1)
                self.toggle_campos_ferias()
                self.ent_aq_de.delete(0, tk.END)
                self.ent_aq_de.insert(0, periodo_ini_usar)
                self.atualizar_fim_aquisitivo()

            if str(apos) in ("1", "Sim", "True"):
                self.var_aposentado.set(1)
            else:
                self.var_aposentado.set(0)

            if str(dom) in ("1", "Sim", "True"):
                self.var_domestica.set(1)
            else:
                self.var_domestica.set(0)
    def formatar_cpf(self, event=None):
        if event and event.keysym in ('BackSpace', 'Delete', 'Left', 'Right'): return
        texto = self.ent_cpf.get()
        numeros = "".join(filter(str.isdigit, texto))[:11]
        resultado = ""
        for i, char in enumerate(numeros):
            if i in [3, 6]: resultado += "."
            elif i == 9: resultado += "-"
            resultado += char
        self.ent_cpf.delete(0, tk.END)
        self.ent_cpf.insert(0, resultado)
        if len(numeros) == 11:
            self.tentar_vincular_historico_funcionario()

    def gerar_termo_govbr(self):
        nome = self.ent_nome.get().strip()
        cpf = self.ent_cpf.get().strip()
        if not nome or not cpf:
            messagebox.showwarning("Aviso", "Preencha o Nome e o CPF do funcionário para emitir o termo.", parent=self.janela_cad)
            return
        arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"Termo_GOVBR_{nome.replace(' ', '_')}.pdf", parent=self.janela_cad)
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
            messagebox.showerror("Erro", f"Erro ao gerar o PDF:\n{str(e)}", parent=self.janela_cad)

    def acionar_busca_cid(self):
        cod = self.ent_cid_cod.get().strip().upper()
        if not cod: return messagebox.showwarning("Aviso", "Digite um código CID para buscar.", parent=self.janela_cad)
        self.ent_cid_cod.delete(0, tk.END)
        self.ent_cid_cod.insert(0, cod)
        resultado = buscar_cid_base(cod)
        self.ent_cid_desc.delete(0, tk.END)
        if resultado: self.ent_cid_desc.insert(0, resultado)
        else: self.ent_cid_desc.focus_set()

    def adicionar_cid_lista(self):
        cod = self.ent_cid_cod.get().strip().upper()
        desc = self.ent_cid_desc.get().strip()
        if not cod: return messagebox.showwarning("Aviso", "Informe o código do CID.", parent=self.janela_cad)
        cod_norm = cod.replace(".", "")
        for item_cod, _ in self.cids_lista:
            if item_cod.replace(".", "") == cod_norm:
                return messagebox.showwarning("Duplicidade", f"O CID '{cod}' já está incluso neste afastamento!", parent=self.janela_cad)
        self.cids_lista.append((cod, desc))
        self.atualizar_tabela_cids()
        self.ent_cid_cod.delete(0, tk.END)
        self.ent_cid_desc.delete(0, tk.END)
        self.ent_cid_cod.focus_set()

    def atualizar_tabela_cids(self):
        for item in self.tree_cid.get_children(): self.tree_cid.delete(item)
        for c in self.cids_lista: self.tree_cid.insert("", "end", values=(c[0], c[1]))

    def excluir_cid_selecionado(self, event):
        selecionado = self.tree_cid.selection()
        if not selecionado: return
        cod_del = self.tree_cid.item(selecionado[0], "values")[0]
        if messagebox.askyesno("Excluir", f"Remover o CID '{cod_del}' da lista?", parent=self.janela_cad):
            self.cids_lista = [c for c in self.cids_lista if c[0] != cod_del]
            self.atualizar_tabela_cids()

    def abrir_modal_editar_cid(self, event):
        selecionado = self.tree_cid.selection()
        if not selecionado: return
        idx = self.tree_cid.index(selecionado[0])
        cod_atual, desc_atual = self.cids_lista[idx]
        jan_ed = tk.Toplevel(self.janela_cad)
        jan_ed.title(f"Editar CID | {cod_atual}")
        jan_ed.geometry("450x180")
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
        def salvar_ed():
            c_novo = ed_cod.get().strip().upper()
            d_novo = ed_desc.get().strip()
            if not c_novo: return messagebox.showwarning("Aviso", "O código não pode ficar vazio.", parent=jan_ed)
            c_norm = c_novo.replace(".", "")
            for i, (item_cod, _) in enumerate(self.cids_lista):
                if i != idx and item_cod.replace(".", "") == c_norm:
                    return messagebox.showwarning("Duplicidade", f"O CID '{c_novo}' já está incluso neste afastamento!", parent=jan_ed)
            self.cids_lista[idx] = (c_novo, d_novo)
            self.atualizar_tabela_cids()
            jan_ed.destroy()
        tk.Button(jan_ed, text="💾 Salvar Alteração", command=salvar_ed, bg="#3B82F6", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", pady=4, padx=15).pack(pady=(0, 10))

    def criar_calendario(self, parent_frame, width=15, cor_fundo='#3B82F6'):
        return DateEntryLivre(parent_frame, width=width, background=cor_fundo, foreground='white', borderwidth=0, date_pattern='dd/mm/yyyy', state="normal", font=("Segoe UI", 10))

    def calcular_por_datas(self, origem=None):
        try:
            de_str = self.ent_per_de.get().strip()
            if len(de_str) == 10:
                data_de = datetime.strptime(de_str, "%d/%m/%Y")
                if origem == "de":
                    self.ent_16dia.set_date(data_de + timedelta(days=15))

                ate_str = self.ent_per_ate.get().strip()
                dias_str = self.ent_dias.get().strip()

                if len(ate_str) == 10:
                    data_ate = datetime.strptime(ate_str, "%d/%m/%Y")
                    dias = (data_ate - data_de).days + 1
                    self.ent_dias.delete(0, tk.END)
                    if dias > 0:
                        self.ent_dias.insert(0, str(dias))
                        self.verificar_obrigatoriedade_ferias(dias)
                    else:
                        self.verificar_obrigatoriedade_ferias(0)
                elif origem == "de" and dias_str.isdigit() and int(dias_str) > 0:
                    dias = int(dias_str)
                    self.ent_per_ate.set_date(data_de + timedelta(days=dias - 1))
                    self.verificar_obrigatoriedade_ferias(dias)
        except ValueError: pass

    def calcular_por_dias(self, event=None):
        try:
            de_str = self.ent_per_de.get().strip()
            dias_str = self.ent_dias.get().strip()
            if len(de_str) == 10 and dias_str.isdigit():
                dias = int(dias_str)
                if dias > 0:
                    data_de = datetime.strptime(de_str, "%d/%m/%Y")
                    self.ent_per_ate.set_date(data_de + timedelta(days=dias - 1))
                    if not self.ent_16dia.get().strip():
                        self.ent_16dia.set_date(data_de + timedelta(days=15))
                    self.verificar_obrigatoriedade_ferias(dias)
            elif not dias_str.isdigit() or int(dias_str) <= 0:
                self.verificar_obrigatoriedade_ferias(0)
        except ValueError: pass

    def salvar_registro(self):
        emp = self.ent_empresa.get().strip()
        cod_empresa = self.ent_cod_empresa.get().strip()
        nome = self.ent_nome.get().strip()
        cod_func = self.ent_cod_func.get().strip()
        cpf_func = self.ent_cpf.get().strip()
        motivo = self.combo_motivo.get()
        ult_dia = self.ent_ult_dia.get().strip()
        obrigacao = self.ent_obrigacao.get().strip()
        per_de = self.ent_per_de.get().strip()
        per_ate = self.ent_per_ate.get().strip()
        dias_str = self.ent_dias.get().strip()
        d16 = self.ent_16dia.get().strip()

        if not nome or not emp:
            messagebox.showwarning("Aviso", "A Empresa e o Nome são campos obrigatórios.", parent=self.janela_cad)
            return

        dias_val = int(dias_str) if dias_str.isdigit() else 0
        if dias_val > 15:
            self.var_ctrl_ferias.set(1)

        ctrl_ferias = self.var_ctrl_ferias.get()
        aq_de = self.ent_aq_de.get().strip() if ctrl_ferias == 1 else ""
        aq_ate = self.ent_aq_ate.get().strip() if ctrl_ferias == 1 else ""

        if ctrl_ferias == 1 and (not aq_de or len(aq_de) != 10):
            messagebox.showwarning("Aviso", "Como o afastamento é superior a 15 dias, é obrigatório preencher o 'Início Per. Aquisitivo' para o controle de Férias.", parent=self.janela_cad)
            self.toggle_campos_ferias()
            return

        cid_pendente = self.ent_cid_cod.get().strip().upper()
        desc_pendente = self.ent_cid_desc.get().strip()
        if cid_pendente:
            pend_norm = cid_pendente.replace(".", "")
            if not any(c[0].replace(".", "") == pend_norm for c in self.cids_lista):
                self.cids_lista.append((cid_pendente, desc_pendente))

        cid_resumo = " | ".join([c[0] for c in self.cids_lista]) if self.cids_lista else ""
        estab_herdada = ""

        aborto_val = self.var_aborto.get() if motivo == "Maternidade" else 0
        dt_nasc_val = self.ent_data_nasc.get().strip() if motivo == "Maternidade" else ""
        sal_fam_val = self.combo_sal_fam.get() if motivo == "Maternidade" else ""
        sal_fam_obs_val = self.ent_obs_sal_fam.get().strip() if motivo == "Maternidade" else ""
        imp_renda_val = self.combo_imp_renda.get() if motivo == "Maternidade" else ""
        imp_renda_obs_val = self.ent_obs_imp_renda.get().strip() if motivo == "Maternidade" else ""

        try:
            conn = conectar_db()
            cursor = conn.cursor()
            
            if motivo == "Maternidade":
                selecao = self.combo_gestante.get()
                if selecao in self.map_gestantes:
                    id_gest = self.map_gestantes[selecao]['id']
                    estab_herdada = self.map_gestantes[selecao]['estab']
                    cursor.execute("UPDATE gestantes SET status='Afastada' WHERE id=?", (id_gest,))

            cursor.execute('''
                INSERT INTO afastamentos (
                    empresa, cod_empresa, cod_empregado, nome_empregado, tipo_afastamento, cid, 
                    ultimo_dia_trab, data_obrigacao, periodo_de, periodo_ate, qtd_dias, decimo_sexto_dia, retornou,
                    aposentado, domestica, estabilidade, cpf,
                    aborto, data_nascimento, salario_familia, salario_familia_obs, imposto_renda, imposto_renda_obs,
                    controlar_ferias, per_aquisitivo_de, per_aquisitivo_ate
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                emp, cod_empresa, cod_func, nome, motivo, cid_resumo, 
                ult_dia, obrigacao, per_de, per_ate, dias_str, d16, 
                self.var_aposentado.get(), self.var_domestica.get(), estab_herdada, cpf_func,
                aborto_val, dt_nasc_val, sal_fam_val, sal_fam_obs_val, imp_renda_val, imp_renda_obs_val,
                ctrl_ferias, aq_de, aq_ate
            ))
            
            id_novo = cursor.lastrowid
            
            if per_de and per_ate:
                cursor.execute("INSERT INTO atestados_historico (id_afastamento, data_de, data_ate, qtd_dias, observacoes) VALUES (?, ?, ?, ?, ?)", (id_novo, per_de, per_ate, dias_str, "Atestado Inicial"))
                               
            for c_cod, c_desc in self.cids_lista:
                cursor.execute("INSERT INTO cids_historico (id_afastamento, codigo, descricao) VALUES (?, ?, ?)", (id_novo, c_cod, c_desc))

            conn.commit()
            conn.close()

            if ctrl_ferias == 1 and aq_de:
                sincronizar_aquisitivo_funcionario(cpf_func, cod_empresa, cod_func, ctrl_ferias, aq_de, aq_ate)

            messagebox.showinfo("Sucesso", "Novo Afastamento registado com sucesso!", parent=self.janela_cad)
            self.callback_atualizar()
            self.janela_cad.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao guardar:\n{str(e)}", parent=self.janela_cad)