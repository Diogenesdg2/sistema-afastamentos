import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import platform
import subprocess
from datetime import datetime, timedelta
from tkcalendar import DateEntry
import sqlite3

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from database import inicializar_db, conectar_db, calcular_dias_beneficio_aquisitivo
from cadastro import TelaCadastro
from importar import executar_importacao
from manutencao import TelaManutencao
from relatorio import gerar_pdf_afastamentos, gerar_pdf_gestantes, gerar_excel_afastamentos, gerar_excel_gestantes
from gestante import TelaGestante
from manutencao_gestante import TelaManutencaoGestante

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

class SistemaAfastamentos:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle de Afastamentos e Gestantes INSS - Viman")
        self.root.geometry("1350x780") 
        self.root.configure(bg="#F3F4F6")

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TFrame", background="#F3F4F6")
        style.configure("TLabelframe", background="#FFFFFF", borderwidth=2)
        style.configure("TLabelframe.Label", background="#FFFFFF", font=("Segoe UI", 10, "bold"), foreground="#374151")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#E5E7EB", foreground="#1F2937")
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=25)
        
        style.configure("TNotebook", background="#F3F4F6", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 11, "bold"), padding=[20, 10], background="#E5E7EB", foreground="#374151")
        style.map("TNotebook.Tab", background=[("selected", "#FFFFFF")], foreground=[("selected", "#1D4ED8")])

        frame_acoes = tk.Frame(self.root, bg="#F3F4F6")
        frame_acoes.pack(fill="x", padx=20, pady=(15, 10))

        btn_novo = tk.Button(frame_acoes, text="➕ Novo Afastamento", command=self.abrir_cadastro, 
                             bg="#10B981", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", padx=15, pady=6)
        btn_novo.pack(side="left", padx=(0, 10))

        btn_gestante = tk.Button(frame_acoes, text="🤰 Nova Gestante", command=lambda: TelaGestante(self.root, self.filtrar_dados), 
                             bg="#DB2777", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", padx=15, pady=6)
        btn_gestante.pack(side="left", padx=(0, 10))

        btn_atualizar = tk.Button(frame_acoes, text="🔄 Atualizar Tudo", command=self.filtrar_dados, 
                                  bg="#3B82F6", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", padx=15, pady=6)
        btn_atualizar.pack(side="left")

        btn_pdf = tk.Button(frame_acoes, text="📄 PDF", command=self.exportar_pdf, 
                            bg="#EC4899", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", padx=15, pady=6)
        btn_pdf.pack(side="right", padx=(10, 0))

        btn_excel = tk.Button(frame_acoes, text="📊 Excel", command=self.exportar_excel, 
                            bg="#16A34A", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", padx=15, pady=6)
        btn_excel.pack(side="right", padx=(10, 0))

        btn_importar = tk.Button(frame_acoes, text="📂 Importar XLSX", command=self.chamar_importacao, 
                                 bg="#6366F1", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", padx=15, pady=6)
        btn_importar.pack(side="right", padx=(10, 0))

        frame_filtros = ttk.LabelFrame(self.root, text=" 🔎 Filtros e Pesquisa Rápida ", padding=10)
        frame_filtros.pack(fill="x", padx=20, pady=(0, 15))

        f_row1 = tk.Frame(frame_filtros, bg="#FFFFFF")
        f_row1.pack(fill="x", pady=(0, 10))

        fonte_lbl = ("Segoe UI", 9, "bold")
        fonte_ent = ("Segoe UI", 9)

        tk.Label(f_row1, text="Nome:", bg="white", font=fonte_lbl).pack(side="left", padx=(5, 2))
        self.ent_busca = tk.Entry(f_row1, width=25, font=fonte_ent, relief="solid", bd=1)
        self.ent_busca.pack(side="left", padx=(0, 15))
        self.ent_busca.bind("<KeyRelease>", self.filtrar_dados)
        
        tk.Label(f_row1, text="Empresa:", bg="white", font=fonte_lbl).pack(side="left", padx=(0, 2))
        self.ent_empresa_filtro = tk.Entry(f_row1, width=20, font=fonte_ent, relief="solid", bd=1)
        self.ent_empresa_filtro.pack(side="left", padx=(0, 15))
        self.ent_empresa_filtro.bind("<KeyRelease>", self.filtrar_dados)

        tk.Label(f_row1, text="Data Obrigação:", bg="white", font=fonte_lbl).pack(side="left", padx=(0, 2))
        self.ent_data_filtro = DateEntryLivre(f_row1, width=12, font=fonte_ent, background='#3B82F6', foreground='white', borderwidth=0, date_pattern='dd/mm/yyyy')
        self.ent_data_filtro.pack(side="left", padx=(0, 5))
        self.ent_data_filtro.delete(0, tk.END) 
        self.ent_data_filtro.bind("<<DateEntrySelected>>", self.filtrar_dados) 
        self.ent_data_filtro.bind("<KeyRelease>", self.filtrar_dados) 
        
        btn_hoje = tk.Button(f_row1, text="Hoje", command=self.filtrar_hoje, bg="#F59E0B", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2")
        btn_hoje.pack(side="left", padx=5)

        btn_limpar = tk.Button(f_row1, text="Limpar Filtros", command=self.limpar_filtros, bg="#9CA3AF", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2")
        btn_limpar.pack(side="left", padx=15)

        f_row2 = tk.Frame(frame_filtros, bg="#FFFFFF")
        f_row2.pack(fill="x")

        tk.Label(f_row2, text="Tipos:", bg="white", font=("Segoe UI", 9, "bold"), fg="#1D4ED8").pack(side="left", padx=(5, 5))
        
        self.var_t_doenca = tk.IntVar(value=1)
        tk.Checkbutton(f_row2, text="Doença", variable=self.var_t_doenca, bg="white", font=fonte_ent, command=self.filtrar_dados).pack(side="left", padx=2)
        self.var_t_acidente = tk.IntVar(value=1)
        tk.Checkbutton(f_row2, text="Acidente", variable=self.var_t_acidente, bg="white", font=fonte_ent, command=self.filtrar_dados).pack(side="left", padx=2)
        self.var_t_carcere = tk.IntVar(value=1)
        tk.Checkbutton(f_row2, text="Cárcere", variable=self.var_t_carcere, bg="white", font=fonte_ent, command=self.filtrar_dados).pack(side="left", padx=2)
        self.var_t_militar = tk.IntVar(value=1)
        tk.Checkbutton(f_row2, text="Militar", variable=self.var_t_militar, bg="white", font=fonte_ent, command=self.filtrar_dados).pack(side="left", padx=2)
        self.var_t_maternidade = tk.IntVar(value=1)
        tk.Checkbutton(f_row2, text="Maternidade", variable=self.var_t_maternidade, bg="white", font=fonte_ent, command=self.filtrar_dados).pack(side="left", padx=2)
        self.var_t_outros = tk.IntVar(value=1)
        tk.Checkbutton(f_row2, text="Outros", variable=self.var_t_outros, bg="white", font=fonte_ent, command=self.filtrar_dados).pack(side="left", padx=2)

        ttk.Separator(f_row2, orient="vertical").pack(side="left", fill="y", padx=20)

        tk.Label(f_row2, text="Status:", bg="white", font=("Segoe UI", 9, "bold"), fg="#1D4ED8").pack(side="left", padx=(5, 5))
        
        self.var_s_ativo = tk.IntVar(value=1)
        tk.Checkbutton(f_row2, text="Ativos (Afastado / Gestante)", variable=self.var_s_ativo, bg="white", font=fonte_ent, fg="#B45309", command=self.filtrar_dados).pack(side="left", padx=2)
        
        self.var_s_inativo = tk.IntVar(value=0)
        tk.Checkbutton(f_row2, text="Baixados (Retornou / Concluído)", variable=self.var_s_inativo, bg="white", font=fonte_ent, fg="#15803D", command=self.filtrar_dados).pack(side="left", padx=2)

        self.paned = ttk.PanedWindow(self.root, orient="vertical")
        self.paned.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.abas_main = ttk.Notebook(self.paned)
        self.paned.add(self.abas_main, weight=3)

        self.aba_afast = ttk.Frame(self.abas_main)
        self.aba_gest = ttk.Frame(self.abas_main)

        self.abas_main.add(self.aba_afast, text=" 🏥 Afastamentos INSS ")
        self.abas_main.add(self.aba_gest, text=" 🤰 Acompanhamento de Gestantes ")
        self.abas_main.bind("<<NotebookTabChanged>>", self.limpar_preview_on_tab_change)

        colunas_afast = ("ID", "Empresa", "Nome", "Tipo", "Últ. Dia", "Obrigação", "Status")
        self.tree = ttk.Treeview(self.aba_afast, columns=colunas_afast, show="headings")
        larg_afast = [40, 150, 250, 80, 80, 100, 170]
        for col, larg in zip(colunas_afast, larg_afast):
            self.tree.heading(col, text=col, command=lambda _col=col: self.ordenar_coluna(_col, False))
            self.tree.column(col, width=larg, anchor="center" if col not in ["Empresa", "Nome"] else "w")

        self.tree.tag_configure('retornou', background='#C8E6C9') 
        self.tree.tag_configure('atrasado', background='#FFCDD2') 
        self.tree.tag_configure('hoje', background='#FFF9C4')     
        self.tree.tag_configure('afastado', background='#FFF3E0') 
        self.tree.tag_configure('tratada', background='#E2E8F0')
        self.tree.tag_configure('alerta_maternidade', background='#F3E8FF', foreground='#7E22CE')
        self.tree.tag_configure('alerta_16dia', background='#FFEDD5', foreground='#C2410C')
        self.tree.tag_configure('alerta_prorrogacao', background='#E0F2FE', foreground='#0284C7')
        # Alterado para azul (Checar Retorno)
        self.tree.tag_configure('alerta_retorno', background='#DBEAFE', foreground='#1E40AF')
        # Nova cor para Aguardando data de retorno (Cárcere / Militar)
        self.tree.tag_configure('aguardando_retorno', background="#9AA09F", foreground="#3B4140")

        scroll_afast = ttk.Scrollbar(self.aba_afast, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_afast.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll_afast.pack(side="right", fill="y")
        
        self.tree.bind("<Double-1>", self.abrir_manutencao)
        self.tree.bind("<<TreeviewSelect>>", self.atualizar_preview)
        self.tree.bind("<Shift-Delete>", self.excluir_funcionario_oculto)

        colunas_gest = ("ID", "Empresa", "Nome", "Status", "Exame", "Semanas", "Prev. Parto (40 Sem.)", "Estabilidade")
        self.tree_gest = ttk.Treeview(self.aba_gest, columns=colunas_gest, show="headings")
        larg_gest = [40, 150, 220, 80, 80, 70, 120, 100]
        for col, larg in zip(colunas_gest, larg_gest):
            self.tree_gest.heading(col, text=col)
            self.tree_gest.column(col, width=larg, anchor="center" if col not in ["Empresa", "Nome"] else "w")

        self.tree_gest.tag_configure('alerta_parto', background='#FCE7F3', foreground='#BE185D') 
        self.tree_gest.tag_configure('concluido', background='#E2E8F0', foreground='#64748B')
        self.tree_gest.tag_configure('afastada', background='#C8E6C9', foreground='#166534')

        scroll_gest = ttk.Scrollbar(self.aba_gest, orient="vertical", command=self.tree_gest.yview)
        self.tree_gest.configure(yscrollcommand=scroll_gest.set)
        self.tree_gest.pack(side="left", fill="both", expand=True)
        scroll_gest.pack(side="right", fill="y")

        self.tree_gest.bind("<Double-1>", self.abrir_manutencao_gestante)
        self.tree_gest.bind("<<TreeviewSelect>>", self.atualizar_preview_gestante)
        self.tree_gest.bind("<Shift-Delete>", self.excluir_gestante)

        frame_preview = ttk.LabelFrame(self.paned, text=" Pré-visualização e Detalhes Rápidos ", padding=5)
        self.paned.add(frame_preview, weight=1)
        
        self.txt_preview = tk.Text(frame_preview, height=11, bg="#FFFFFF", font=("Consolas", 10), state="disabled", wrap="word", relief="flat")
        self.txt_preview.pack(side="left", fill="both", expand=True, padx=(5, 5), pady=5)

        self.txt_historico = tk.Text(frame_preview, height=11, bg="#F0F8FF", font=("Consolas", 10), state="disabled", wrap="word", width=48, relief="flat")
        self.txt_historico.pack(side="right", fill="y", padx=(0, 5), pady=5)

        self.corrigir_datas_antigas_gestantes()
        
        self.carregar_dados()
        self.root.after(500, self.verificar_lembretes_startup)

    def corrigir_datas_antigas_gestantes(self):
        conn = conectar_db()
        cursor = conn.cursor()
        cursor.execute("SELECT id, data_exame, semanas_exame FROM gestantes")
        for linha in cursor.fetchall():
            id_g, dt_exame_str, semanas_val = linha
            if dt_exame_str and semanas_val is not None and str(semanas_val).strip() != "":
                try:
                    dt_exame = datetime.strptime(dt_exame_str, "%d/%m/%Y")
                    semanas = float(str(semanas_val).replace(",", "."))
                    dias_para_40 = (40.0 - semanas) * 7
                    dt_40 = dt_exame + timedelta(days=dias_para_40)
                    nova_data = dt_40.strftime("%d/%m/%Y")
                    cursor.execute("UPDATE gestantes SET data_37_semanas=? WHERE id=?", (nova_data, id_g))
                except Exception:
                    pass
        conn.commit()
        conn.close()

    def exportar_lembretes_pdf(self, alertas):
        if not alertas: return
        arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Documento PDF", "*.pdf")], initialfile=f"Lembretes_RH_{datetime.now().strftime('%d%m%Y')}.pdf")
        if not arquivo: return

        doc = SimpleDocTemplate(arquivo, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        elementos = []
        estilos = getSampleStyleSheet()
        estilo_titulo = ParagraphStyle('Titulo', parent=estilos['Heading1'], fontSize=16, textColor=colors.HexColor("#DC2626"), alignment=1, spaceAfter=20)
        estilo_texto = estilos['Normal']

        elementos.append(Paragraph("Resumo Matinal de Lembretes Pendentes", estilo_titulo))

        dados_tabela = [["Tipo de Alerta", "Colaborador / Detalhe do Aviso"]]
        for al in alertas:
            if al['tipo'] == 'mat': tipo_str = "Fim Licença Maternidade"
            elif al['tipo'] == 'inss': tipo_str = "Análise Documental"
            elif al['tipo'] == 'prorrog': tipo_str = "Prorrogação de Benefício"
            elif al['tipo'] == 'nasc': tipo_str = "Verificação de Nascimento"
            elif al['tipo'] == 'retorno': tipo_str = "Verificação de Retorno"
            elif al['tipo'] == 'ferias': tipo_str = "Perda de Férias (Art. 133)"
            else: tipo_str = "Aviso"
            
            msg_limpa = al['msg'].replace("🤱", "").replace("🚨", "").replace("📅", "").replace("⚠️", "").replace("⏳", "").replace("👶", "").replace("✅", "").replace("🏖️", "").strip()
            dados_tabela.append([tipo_str, Paragraph(msg_limpa, estilo_texto)])

        tabela = Table(dados_tabela, colWidths=[130, 400])
        tabela.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E5E7EB")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.black),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('BOX', (0,0), (-1,-1), 1, colors.grey),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
        ]))

        elementos.append(tabela)

        try:
            doc.build(elementos)
            if platform.system() == 'Windows': os.startfile(arquivo)
            elif platform.system() == 'Darwin': subprocess.call(('open', arquivo))
            else: subprocess.call(('xdg-open', arquivo))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar o PDF:\n{str(e)}")

    def verificar_lembretes_startup(self):
        alertas = []
        conn = conectar_db()
        cursor = conn.cursor()
        hoje = datetime.now().date()
        
        # 1. MATERNIDADE (LICENÇA)
        cursor.execute("SELECT id, nome_empregado, periodo_ate FROM afastamentos WHERE tipo_afastamento='Maternidade' AND retornou=0 AND (alerta_mat_tratado=0 OR alerta_mat_tratado IS NULL)")
        for linha in cursor.fetchall():
            id_af, nome, per_ate_str = linha
            if per_ate_str:
                try:
                    dt_ate = datetime.strptime(per_ate_str, "%d/%m/%Y").date()
                    dias_restantes = (dt_ate - hoje).days
                    if 0 <= dias_restantes <= 7: alertas.append({'id': id_af, 'tipo': 'mat', 'msg': f"🤱 {nome}: Licença termina em {dias_restantes} dias", 'cor': '#7E22CE'})
                    elif dias_restantes < 0: alertas.append({'id': id_af, 'tipo': 'mat', 'msg': f"🚨 {nome}: Licença VENCEU há {abs(dias_restantes)} dias!", 'cor': '#DC2626'})
                except ValueError: pass

        # 2. GESTANTES (PARTO)
        cursor.execute("SELECT id, nome_empregado, data_37_semanas FROM gestantes WHERE status='Gestante' AND alerta_nasc_tratado=0")
        for linha in cursor.fetchall():
            id_g, nome, dt_40_str = linha
            if dt_40_str:
                try:
                    dt_40 = datetime.strptime(dt_40_str, "%d/%m/%Y").date()
                    dias_para_40 = (dt_40 - hoje).days
                    if dias_para_40 <= 21:
                        if dias_para_40 >= 0: alertas.append({'id': id_g, 'tipo': 'nasc', 'msg': f"👶 {nome}: Entrou na 37ª semana! (Prev. Parto em {dias_para_40}d)", 'cor': '#DB2777'})
                        else: alertas.append({'id': id_g, 'tipo': 'nasc', 'msg': f"🚨 {nome}: Passou da 40ª semana há {abs(dias_para_40)} dias. Já nasceu?", 'cor': '#DC2626'})
                except ValueError: pass

        # 3. 16º DIA (INSS - Ignora Aposentados, Cárcere, Militar e Maternidade)
        cursor.execute("SELECT id, nome_empregado, decimo_sexto_dia, qtd_dias, aposentado FROM afastamentos WHERE retornou=0 AND tipo_afastamento NOT IN ('Cárcere', 'Militar', 'Maternidade') AND decimo_sexto_dia IS NOT NULL AND decimo_sexto_dia != '' AND (alerta_16_tratado=0 OR alerta_16_tratado IS NULL)")
        for linha in cursor.fetchall():
            id_af, nome, d16_str, qtd_dias_str, aposentado_val = linha
            if str(aposentado_val) in ("1", "Sim", "True"): continue

            try:
                qtd_dias = int(qtd_dias_str)
                if qtd_dias <= 15: continue
            except (ValueError, TypeError): pass

            try:
                dt_16 = datetime.strptime(d16_str, "%d/%m/%Y").date()
                if hoje >= dt_16:
                    c2 = conn.cursor()
                    c2.execute("SELECT count(*) FROM analises_inss WHERE id_afastamento=?", (id_af,))
                    if c2.fetchone()[0] == 0:
                        c2.execute("SELECT count(*) FROM pericias_inss WHERE id_afastamento=?", (id_af,))
                        if c2.fetchone()[0] == 0:
                            dias_atraso = (hoje - dt_16).days
                            if dias_atraso == 0: alertas.append({'id': id_af, 'tipo': 'inss', 'msg': f"📅 {nome}: 16º Dia HOJE! Checar Análise Documental.", 'cor': '#D97706'})
                            else: alertas.append({'id': id_af, 'tipo': 'inss', 'msg': f"⚠️ {nome}: 16º Dia passou há {dias_atraso} dias. Checar Análise Doc!", 'cor': '#DC2626'})
            except ValueError: pass

        # 4. PRORROGAÇÃO (INSS)
        cursor.execute("""
            SELECT an.id, af.nome_empregado, an.concessao_de, an.concessao_ate 
            FROM analises_inss an
            JOIN afastamentos af ON an.id_afastamento = af.id
            WHERE af.retornou=0 AND af.tipo_afastamento NOT IN ('Maternidade', 'Cárcere', 'Militar') 
            AND an.situacao = 'Concedido' AND (an.prorrogacao_tratada=0 OR an.prorrogacao_tratada IS NULL)
        """)
        for linha in cursor.fetchall():
            id_an, nome, c_de, c_ate = linha
            if c_de and c_ate:
                try:
                    dt_de = datetime.strptime(c_de, "%d/%m/%Y").date()
                    dt_ate = datetime.strptime(c_ate, "%d/%m/%Y").date()
                    if (dt_ate - dt_de).days > 15:
                        dias_para_fim = (dt_ate - hoje).days
                        if dias_para_fim <= 15:
                            if dias_para_fim >= 0:
                                alertas.append({'id': id_an, 'tipo': 'prorrog', 'msg': f"⏳ {nome}: Checar Prorrogação (Análise Documental #{id_an} vence em {dias_para_fim} dias)", 'cor': '#0284C7'})
                            else:
                                alertas.append({'id': id_an, 'tipo': 'prorrog', 'msg': f"🚨 {nome}: Prorrogação Atrasada! (Análise #{id_an} venceu há {abs(dias_para_fim)} dias)", 'cor': '#DC2626'})
                except ValueError: pass

        # 5. RETORNO AO TRABALHO (Ignora Cárcere sem retorno definido)
        cursor.execute("SELECT id, nome_empregado, periodo_ate FROM afastamentos WHERE retornou=0 AND tipo_afastamento != 'Cárcere' AND periodo_ate IS NOT NULL AND periodo_ate != '' AND (alerta_retorno_tratado=0 OR alerta_retorno_tratado IS NULL)")
        for linha in cursor.fetchall():
            id_af, nome, p_ate_str = linha
            try:
                dt_ate = datetime.strptime(p_ate_str, "%d/%m/%Y").date()
                dias_pos = (hoje - dt_ate).days
                if dias_pos >= 1:
                    alertas.append({'id': id_af, 'tipo': 'retorno', 'msg': f"✅ {nome}: Fim do atestado foi há {dias_pos} dias. Já retornou?", 'cor': '#1E40AF'})
            except ValueError: pass

        # 6. ALERTA DE PERDA DE FÉRIAS (180 DIAS)
        cursor.execute("SELECT id, cpf, cod_empresa, cod_empregado, nome_empregado, per_aquisitivo_de, per_aquisitivo_ate FROM afastamentos WHERE retornou=0 AND controlar_ferias=1 AND (alerta_ferias_tratado=0 OR alerta_ferias_tratado IS NULL)")
        for linha in cursor.fetchall():
            id_af, cpf, cod_emp, cod_func, nome, aq_de, aq_ate = linha
            if aq_de and aq_ate:
                dias_benef, _ = calcular_dias_beneficio_aquisitivo(cpf, cod_emp, cod_func, aq_de, aq_ate, id_atual=id_af)
                if dias_benef >= 180:
                    alertas.append({'id': id_af, 'tipo': 'ferias', 'msg': f"🚨 {nome}: Atingiu {dias_benef} dias de benefício! Direito a férias perdido (Art. 133, IV).", 'cor': '#DC2626'})
                elif dias_benef >= 150:
                    alertas.append({'id': id_af, 'tipo': 'ferias', 'msg': f"⚠️ {nome}: {dias_benef}/180 dias de benefício concedido. Risco de perder férias!", 'cor': '#D97706'})

        conn.close()

        if alertas:
            jan_alerta = tk.Toplevel(self.root)
            jan_alerta.title("🔔 Resumo Matinal de Lembretes")
            jan_alerta.geometry("850x580") 
            jan_alerta.configure(bg="#FFFFFF")
            
            jan_alerta.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 425
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 290
            jan_alerta.geometry(f"+{x}+{y}")
            
            tk.Label(jan_alerta, text="Atenção! Lembretes Pendentes:", font=("Segoe UI", 16, "bold"), bg="white", fg="#EF4444").pack(pady=(20, 10))
            
            f_lista = tk.Frame(jan_alerta, bg="white")
            f_lista.pack(fill="both", expand=True, padx=20)
            
            canvas = tk.Canvas(f_lista, bg="white", highlightthickness=0)
            scroll = ttk.Scrollbar(f_lista, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg="white")

            scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scroll.set)

            canvas.pack(side="left", fill="both", expand=True)
            if len(alertas) > 6: scroll.pack(side="right", fill="y")

            def cmd_tratar(id_alvo, tipo_alvo, frame_alvo):
                c_upd = conectar_db()
                if tipo_alvo == 'inss': c_upd.cursor().execute("UPDATE afastamentos SET alerta_16_tratado=1 WHERE id=?", (id_alvo,))
                elif tipo_alvo == 'mat': c_upd.cursor().execute("UPDATE afastamentos SET alerta_mat_tratado=1 WHERE id=?", (id_alvo,))
                elif tipo_alvo == 'prorrog': c_upd.cursor().execute("UPDATE analises_inss SET prorrogacao_tratada=1 WHERE id=?", (id_alvo,))
                elif tipo_alvo == 'nasc': c_upd.cursor().execute("UPDATE gestantes SET alerta_nasc_tratado=1 WHERE id=?", (id_alvo,))
                elif tipo_alvo == 'retorno': c_upd.cursor().execute("UPDATE afastamentos SET alerta_retorno_tratado=1 WHERE id=?", (id_alvo,))
                elif tipo_alvo == 'ferias': c_upd.cursor().execute("UPDATE afastamentos SET alerta_ferias_tratado=1 WHERE id=?", (id_alvo,))
                c_upd.commit()
                c_upd.close()
                frame_alvo.destroy() 
                self.carregar_dados() 

            for al in alertas: 
                f_item = tk.Frame(scrollable_frame, bg="#F9FAFB", bd=1, relief="solid")
                f_item.pack(fill="x", pady=4, padx=5)
                
                lbl = tk.Label(f_item, text=al['msg'], font=("Segoe UI", 12, "bold"), fg=al['cor'], bg="#F9FAFB", wraplength=650, justify="left")
                lbl.pack(side="left", padx=10, pady=12)
                
                btn = tk.Button(f_item, text="✔ Marcar Feito", bg="#10B981", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", command=lambda i=al['id'], t=al['tipo'], fr=f_item: cmd_tratar(i, t, fr))
                btn.pack(side="right", padx=15, pady=12)
                
            f_botoes = tk.Frame(jan_alerta, bg="white")
            f_botoes.pack(pady=20)
            
            btn_pdf = tk.Button(f_botoes, text="📄 Exportar Lembretes (PDF)", command=lambda: self.exportar_lembretes_pdf(alertas), bg="#EC4899", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", pady=8, padx=20)
            btn_pdf.pack(side="left", padx=10)
            
            btn_ok = tk.Button(f_botoes, text="Ok, Entendido", command=jan_alerta.destroy, bg="#3B82F6", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", pady=8, padx=25)
            btn_ok.pack(side="left", padx=10)

    def excluir_funcionario_oculto(self, event):
        selecionado = self.tree.selection()
        if not selecionado: return
        
        valores = self.tree.item(selecionado[0], "values")
        id_registro = valores[0]
        nome_func = valores[2]
        
        if messagebox.askyesno("Exclusão Oculta", f"ATENÇÃO: Deseja apagar '{nome_func}' (ID: {id_registro}) e TODO o seu histórico?", parent=self.root):
            try:
                conn = conectar_db()
                cursor = conn.cursor()
                cursor.execute("DELETE FROM atestados_historico WHERE id_afastamento=?", (id_registro,))
                cursor.execute("DELETE FROM cids_historico WHERE id_afastamento=?", (id_registro,))
                cursor.execute("DELETE FROM analises_inss WHERE id_afastamento=?", (id_registro,))
                cursor.execute("DELETE FROM pericias_inss WHERE id_afastamento=?", (id_registro,))
                cursor.execute("DELETE FROM afastamentos WHERE id=?", (id_registro,))
                conn.commit()
                conn.close()
                self.carregar_dados()
            except Exception as e: messagebox.showerror("Erro", f"Erro ao apagar:\n{str(e)}", parent=self.root)

    def excluir_gestante(self, event):
        selecionado = self.tree_gest.selection()
        if not selecionado: return
        
        valores = self.tree_gest.item(selecionado[0], "values")
        id_registro = valores[0]
        nome_func = valores[2]
        
        if messagebox.askyesno("Exclusão", f"Deseja apagar o acompanhamento da gestante '{nome_func}'?", parent=self.root):
            try:
                conn = conectar_db()
                conn.cursor().execute("DELETE FROM gestantes WHERE id=?", (id_registro,))
                conn.commit()
                conn.close()
                self.carregar_dados()
            except Exception as e: messagebox.showerror("Erro", f"Erro ao apagar:\n{str(e)}", parent=self.root)

    def exportar_pdf(self):
        aba_ativa = self.abas_main.index(self.abas_main.select())
        if aba_ativa == 0: 
            ids_tabela = []
            for child in self.tree.get_children(): ids_tabela.append(self.tree.item(child, "values")[0]) 
            if not ids_tabela: return messagebox.showwarning("Aviso", "A tabela de Afastamentos está vazia.")
            arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Documento PDF", "*.pdf")], initialfile=f"Relatorio_Afastamentos_{datetime.now().strftime('%d%m%Y')}.pdf")
            if arquivo: gerar_pdf_afastamentos(ids_tabela, arquivo)
        elif aba_ativa == 1: 
            ids_tabela = []
            for child in self.tree_gest.get_children(): ids_tabela.append(self.tree_gest.item(child, "values")[0]) 
            if not ids_tabela: return messagebox.showwarning("Aviso", "A tabela de Gestantes está vazia.")
            arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Documento PDF", "*.pdf")], initialfile=f"Relatorio_Gestantes_{datetime.now().strftime('%d%m%Y')}.pdf")
            if arquivo: gerar_pdf_gestantes(ids_tabela, arquivo)

    def exportar_excel(self):
        aba_ativa = self.abas_main.index(self.abas_main.select())
        if aba_ativa == 0: 
            ids_tabela = []
            for child in self.tree.get_children(): ids_tabela.append(self.tree.item(child, "values")[0]) 
            if not ids_tabela: return messagebox.showwarning("Aviso", "A tabela de Afastamentos está vazia.")
            arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Documento Excel", "*.xlsx")], initialfile=f"Relatorio_Afastamentos_{datetime.now().strftime('%d%m%Y')}.xlsx")
            if arquivo: gerar_excel_afastamentos(ids_tabela, arquivo)
        elif aba_ativa == 1: 
            ids_tabela = []
            for child in self.tree_gest.get_children(): ids_tabela.append(self.tree_gest.item(child, "values")[0]) 
            if not ids_tabela: return messagebox.showwarning("Aviso", "A tabela de Gestantes está vazia.")
            arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Documento Excel", "*.xlsx")], initialfile=f"Relatorio_Gestantes_{datetime.now().strftime('%d%m%Y')}.xlsx")
            if arquivo: gerar_excel_gestantes(ids_tabela, arquivo)

    def filtrar_hoje(self):
        self.ent_data_filtro.set_date(datetime.now())
        self.var_s_ativo.set(1)
        self.var_s_inativo.set(0)
        self.filtrar_dados()

    def limpar_filtros(self):
        self.ent_busca.delete(0, tk.END)
        self.ent_empresa_filtro.delete(0, tk.END)
        self.ent_data_filtro.delete(0, tk.END)
        
        self.var_t_doenca.set(1)
        self.var_t_acidente.set(1)
        self.var_t_carcere.set(1)
        self.var_t_militar.set(1)
        self.var_t_maternidade.set(1)
        self.var_t_outros.set(1)
        
        self.var_s_ativo.set(1)
        self.var_s_inativo.set(0) 
        self.filtrar_dados()

    def filtrar_dados(self, event=None):
        self.carregar_dados()

    def carregar_dados(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for item in self.tree_gest.get_children(): self.tree_gest.delete(item)
            
        busca = self.ent_busca.get().strip()
        empresa_filtro = self.ent_empresa_filtro.get().strip()
        data_filtro = self.ent_data_filtro.get().strip()

        conn = conectar_db()
        cursor = conn.cursor()
        
        tipos_sel = []
        if self.var_t_doenca.get(): tipos_sel.append("Doença")
        if self.var_t_acidente.get(): tipos_sel.append("Acidente")
        if self.var_t_carcere.get(): tipos_sel.append("Cárcere")
        if self.var_t_militar.get(): tipos_sel.append("Militar")
        if self.var_t_maternidade.get(): tipos_sel.append("Maternidade")
        if self.var_t_outros.get(): tipos_sel.append("Outros")

        status_afast_sel = []
        if self.var_s_ativo.get(): status_afast_sel.append(0) 
        if self.var_s_inativo.get(): status_afast_sel.append(1) 

        hoje_obj = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

        if tipos_sel and status_afast_sel:
            query_afast = "SELECT id, empresa, nome_empregado, tipo_afastamento, ultimo_dia_trab, data_obrigacao, retornou, obrigacao_concluida, periodo_ate, decimo_sexto_dia, qtd_dias, alerta_16_tratado, alerta_mat_tratado, alerta_retorno_tratado, aposentado FROM afastamentos WHERE 1=1"
            params_afast = []
            
            if busca: query_afast += " AND nome_empregado LIKE ?"; params_afast.append(f'%{busca}%')
            if empresa_filtro: query_afast += " AND empresa LIKE ?"; params_afast.append(f'%{empresa_filtro}%')
            if data_filtro: query_afast += " AND data_obrigacao LIKE ?"; params_afast.append(f'%{data_filtro}%')
            
            qmarks_t = ",".join("?" for _ in tipos_sel)
            query_afast += f" AND tipo_afastamento IN ({qmarks_t})"
            params_afast.extend(tipos_sel)

            qmarks_s = ",".join("?" for _ in status_afast_sel)
            query_afast += f" AND retornou IN ({qmarks_s})"
            params_afast.extend(status_afast_sel)

            query_afast += " ORDER BY id DESC"
            cursor.execute(query_afast, params_afast)
            
            for linha in cursor.fetchall():
                id_af, empresa, nome, tipo_afastamento, ult_dia, data_obrig_str, retornou, obrig_tratada, periodo_ate_str, decimo_sexto_dia, qtd_dias_str, alerta_16_trat_val, alerta_mat_trat_val, alerta_retorno_trat_val, aposentado_val = linha
                
                status_texto = "⚠ Afastado"
                tag = 'afastado' 
                
                tem_alerta_maternidade = False
                msg_alerta_mat = ""
                if tipo_afastamento == "Maternidade" and retornou == 0 and periodo_ate_str and not alerta_mat_trat_val:
                    try:
                        dt_ate = datetime.strptime(periodo_ate_str, "%d/%m/%Y").date()
                        dias_diff = (dt_ate - hoje_obj.date()).days
                        if 0 <= dias_diff <= 7:
                            tem_alerta_maternidade = True
                            msg_alerta_mat = f"🚨 Vence em {dias_diff} dias!"
                        elif dias_diff < 0:
                            tem_alerta_maternidade = True
                            msg_alerta_mat = f"🚨 Venceu há {abs(dias_diff)} dias!"
                    except ValueError: pass

                tem_alerta_16 = False
                msg_alerta_16 = ""
                try: qtd_dias = int(qtd_dias_str)
                except (ValueError, TypeError): qtd_dias = 999 

                # Cárcere, Militar e Maternidade não possuem Análise Documental do 16º dia
                if retornou == 0 and tipo_afastamento not in ("Cárcere", "Militar", "Maternidade") and decimo_sexto_dia and qtd_dias > 15 and not alerta_16_trat_val and str(aposentado_val) not in ("1", "Sim", "True"):
                    try:
                        dt_16 = datetime.strptime(decimo_sexto_dia, "%d/%m/%Y").date()
                        if hoje_obj.date() >= dt_16:
                            c2 = conn.cursor()
                            c2.execute("SELECT count(*) FROM analises_inss WHERE id_afastamento=?", (id_af,))
                            if c2.fetchone()[0] == 0:
                                c2.execute("SELECT count(*) FROM pericias_inss WHERE id_afastamento=?", (id_af,))
                                if c2.fetchone()[0] == 0:
                                    tem_alerta_16 = True
                                    dias_pass = (hoje_obj.date() - dt_16).days
                                    if dias_pass == 0: msg_alerta_16 = "📅 Checar Análise Doc."
                                    else: msg_alerta_16 = f"⚠️ Análise Doc. Atrasada ({dias_pass}d)"
                    except ValueError: pass

                tem_alerta_prorrog = False
                msg_alerta_prorrog = ""
                if retornou == 0 and tipo_afastamento not in ('Maternidade', 'Cárcere', 'Militar'):
                    c3 = conn.cursor()
                    c3.execute("SELECT concessao_de, concessao_ate FROM analises_inss WHERE id_afastamento=? AND situacao='Concedido' AND (prorrogacao_tratada=0 OR prorrogacao_tratada IS NULL)", (id_af,))
                    for an_linha in c3.fetchall():
                        if an_linha[0] and an_linha[1]:
                            try:
                                d_de = datetime.strptime(an_linha[0], "%d/%m/%Y").date()
                                d_ate = datetime.strptime(an_linha[1], "%d/%m/%Y").date()
                                if (d_ate - d_de).days > 15:
                                    d_fim = (d_ate - hoje_obj.date()).days
                                    if d_fim <= 15:
                                        tem_alerta_prorrog = True
                                        if d_fim >= 0: msg_alerta_prorrog = f"⏳ Checar Prorrog. ({d_fim}d)"
                                        else: msg_alerta_prorrog = f"🚨 Prorrog. Atrasada ({abs(d_fim)}d)"
                                        break 
                            except ValueError: pass

                tem_alerta_retorno = False
                msg_alerta_retorno = ""
                if retornou == 0 and tipo_afastamento != "Cárcere" and periodo_ate_str and not alerta_retorno_trat_val:
                    try:
                        dt_ate = datetime.strptime(periodo_ate_str, "%d/%m/%Y").date()
                        dias_pos = (hoje_obj.date() - dt_ate).days
                        if dias_pos >= 1:
                            tem_alerta_retorno = True
                            msg_alerta_retorno = f"✅ Checar Retorno ({dias_pos}d)"
                    except ValueError: pass

                if retornou == 1:
                    status_texto = "✔ Retornou"
                    tag = 'retornou' 
                elif tipo_afastamento in ("Cárcere", "Militar"):
                    status_texto = "⏳ Aguardando data de retorno"
                    tag = 'aguardando_retorno'
                elif tem_alerta_retorno: 
                    status_texto = msg_alerta_retorno
                    tag = 'alerta_retorno'
                elif tem_alerta_maternidade:
                    status_texto = msg_alerta_mat
                    tag = 'alerta_maternidade'
                elif tem_alerta_prorrog: 
                    status_texto = msg_alerta_prorrog
                    tag = 'alerta_prorrogacao'
                elif tem_alerta_16:
                    status_texto = msg_alerta_16
                    tag = 'alerta_16dia'
                elif obrig_tratada == 1:
                    status_texto = "✔ Tratada"
                    tag = 'tratada'
                elif data_obrig_str:
                    try:
                        data_obj = datetime.strptime(data_obrig_str, "%d/%m/%Y")
                        if data_obj < hoje_obj: tag = 'atrasado' 
                        elif data_obj == hoje_obj: tag = 'hoje'     
                    except ValueError: pass 

                self.tree.insert("", "end", values=(id_af, empresa, nome, tipo_afastamento, ult_dia, data_obrig_str, status_texto), tags=(tag,))

        # 2. CARREGA GESTANTES
        if "Maternidade" in tipos_sel:
            status_g_sel = []
            if self.var_s_ativo.get(): status_g_sel.extend(["Gestante", "Afastada"])
            if self.var_s_inativo.get(): status_g_sel.append("Concluído")
            
            if status_g_sel:
                query_gest = "SELECT id, empresa, nome_empregado, status, data_exame, semanas_exame, data_37_semanas, estabilidade, alerta_nasc_tratado FROM gestantes WHERE 1=1"
                params_gest = []
                
                if busca: query_gest += " AND nome_empregado LIKE ?"; params_gest.append(f'%{busca}%')
                if empresa_filtro: query_gest += " AND empresa LIKE ?"; params_gest.append(f'%{empresa_filtro}%')
                
                qmarks_sg = ",".join("?" for _ in status_g_sel)
                query_gest += f" AND status IN ({qmarks_sg})"
                params_gest.extend(status_g_sel)
                
                query_gest += " ORDER BY id DESC"
                
                cursor.execute(query_gest, params_gest)
                for linha in cursor.fetchall():
                    tag_g = 'gestante'
                    status_g = linha[3]
                    if status_g == 'Concluído': tag_g = 'concluido'
                    elif status_g == 'Afastada': tag_g = 'afastada'
                    else:
                        dt_40 = linha[6]
                        if dt_40 and not linha[8]:
                            try:
                                d_40 = datetime.strptime(dt_40, "%d/%m/%Y").date()
                                if (d_40 - hoje_obj.date()).days <= 21:
                                    tag_g = 'alerta_parto'
                            except ValueError: pass

                    self.tree_gest.insert("", "end", values=(linha[0], linha[1], linha[2], status_g, linha[4], linha[5], linha[6], linha[7]), tags=(tag_g,))

        conn.close()
        self.limpar_preview()

    def limpar_preview(self):
        self.txt_preview.config(state="normal")
        self.txt_preview.delete("1.0", tk.END)
        self.txt_preview.insert("1.0", "\n  Selecione um registo na tabela para visualizar os detalhes.")
        self.txt_preview.config(state="disabled")

        self.txt_historico.config(state="normal")
        self.txt_historico.delete("1.0", tk.END)
        self.txt_historico.config(state="disabled")

    def limpar_preview_on_tab_change(self, event):
        self.limpar_preview()

    def atualizar_preview(self, event):
        selecionado = self.tree.selection()
        if not selecionado: return self.limpar_preview()
            
        valores = self.tree.item(selecionado[0], "values")
        id_registro = valores[0]

        self.txt_preview.config(state="normal")
        self.txt_preview.delete("1.0", tk.END)
        self.txt_historico.config(state="normal")
        self.txt_historico.delete("1.0", tk.END)

        conn = conectar_db()
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM afastamentos WHERE id=?", (id_registro,))
        linha = cursor.fetchone()
        
        cursor.execute("SELECT codigo, descricao FROM cids_historico WHERE id_afastamento=? ORDER BY id ASC", (id_registro,))
        cids = cursor.fetchall()
        
        cursor.execute("SELECT data_de, data_ate, qtd_dias FROM atestados_historico WHERE id_afastamento=? ORDER BY id ASC", (id_registro,))
        historico = cursor.fetchall()
        
        cursor.execute("SELECT id, data_analise, gerou_pericia, numero_beneficio, situacao, concessao_de, concessao_ate FROM analises_inss WHERE id_afastamento=? ORDER BY id ASC", (id_registro,))
        analises = cursor.fetchall()
        
        cursor.execute("SELECT id_analise_mae, data_agendamento, data_atendimento, especie, numero_beneficio, situacao, concessao_de, concessao_ate FROM pericias_inss WHERE id_afastamento=? ORDER BY id ASC", (id_registro,))
        pericias = cursor.fetchall()
        conn.close()

        if linha:
            obs = linha['observacoes'] if 'observacoes' in linha.keys() and str(linha['observacoes']).strip() != "" else "-"
            aposentado_str = "Sim" if str(linha['aposentado']) in ("1", "Sim", "True") else "Não"
            domestica_str = "Sim" if 'domestica' in linha.keys() and str(linha['domestica']) in ("1", "Sim", "True") else "Não"
            
            cod_empresa_str = linha['cod_empresa'] if 'cod_empresa' in linha.keys() and linha['cod_empresa'] else "-"
            cod_func_str = linha['cod_empregado'] if linha['cod_empregado'] else "-"

            if cids: str_cids = " | ".join([f"{c[0]}" for c in cids])
            else: str_cids = linha['cid'] or "-"

            texto = f"👤 1. DADOS BÁSICOS E AFASTAMENTO\n"
            texto += f"   Empresa: {linha['empresa'] or '-'} (Cód: {cod_empresa_str}) | Funcionário: {linha['nome_empregado'] or '-'} (Cód: {cod_func_str})\n"
            texto += f"   Motivo: {linha['tipo_afastamento'] or '-'} | CID(s): {str_cids} | Aposentado: {aposentado_str} | Doméstica: {domestica_str}\n"
            texto += f"   Início: {linha['periodo_de'] or '-'} | Data Final Prevista: {linha['periodo_ate'] or '-'} | Total de Dias: {linha['qtd_dias'] or '-'}\n"
            texto += f"   Observações Gerais: {obs}\n\n"
            
            tratada = "Sim" if linha['obrigacao_concluida'] else "Não"
            texto += f"⚖️ 2. ACOMPANHAMENTO\n"
            texto += f"   Data Obrigação: {linha['data_obrigacao'] or '-'} (Tratada: {tratada}) | 16º Dia: {linha['decimo_sexto_dia'] or '-'}\n\n"

            gov_auth = linha['gov_br_autorizacao'] if 'gov_br_autorizacao' in linha.keys() and linha['gov_br_autorizacao'] else "-"
            cpf_val = linha['cpf'] if 'cpf' in linha.keys() and linha['cpf'] else "-"
            
            # --- Controle de Férias no Preview ---
            ferias_info = "Desativado"
            if 'controlar_ferias' in linha.keys() and linha['controlar_ferias'] == 1 and linha['per_aquisitivo_de']:
                dias_f, _ = calcular_dias_beneficio_aquisitivo(cpf_val, linha['cod_empresa'], linha['cod_empregado'], linha['per_aquisitivo_de'], linha['per_aquisitivo_ate'], linha['id'])
                ferias_info = f"{dias_f} / 180 dias"

            texto += f"🏛️ 3. INSS E FÉRIAS\n"
            texto += f"   Gov.br: {gov_auth} | CPF: {cpf_val} | Per. Aquisitivo: {ferias_info}\n"
            
            if analises:
                texto += f"   [Análises Documentais]:\n"
                for i, an in enumerate(analises, 1):
                    texto += f"     Nº{i} (ID:{an[0]}): Benefício: {an[3]} | Gerou Per? {an[2]} | Sit: {an[4]} (De: {an[5] or '-'} Até: {an[6] or '-'})\n"
            else: texto += f"   [Análises Documentais]: Nenhuma registada.\n"

            if pericias:
                texto += f"   [Perícias e Concessões]:\n"
                for i, per in enumerate(pericias, 1):
                    texto += f"     Nº{i}: Atend: {per[2]} | Situação: {per[5]} (Benefício: {per[4]})\n"
            else: texto += f"   [Perícias]: Nenhuma registada.\n"

            retornou = "Sim" if linha['retornou'] else "Não (Afastado)"
            hiscre = linha['hiscre'] if 'hiscre' in linha.keys() and linha['hiscre'] else "-"
            
            texto += f"\n🔄 4. RETORNO\n"
            texto += f"   Retornou: {retornou} | Data de Retorno: {linha['retorno_data'] or '-'} | Fim da Estabilidade: {linha['estabilidade'] or '-'} | HISCRE: {hiscre}"
            
            self.txt_preview.insert("1.0", texto)

            texto_hist = "📅 HISTÓRICO DE ATESTADOS\n"
            texto_hist += "----------------------------------------------\n"
            if historico:
                for i, hist in enumerate(historico, start=1): texto_hist += f" {i}º Atestado: De {hist[0]} até {hist[1]} ({hist[2]} dias)\n"
            else: texto_hist += "\n * Nenhum atestado registado."
            self.txt_historico.insert("1.0", texto_hist)

        self.txt_preview.config(state="disabled")
        self.txt_historico.config(state="disabled")

    def atualizar_preview_gestante(self, event):
        selecionado = self.tree_gest.selection()
        if not selecionado: return self.limpar_preview()
            
        valores = self.tree_gest.item(selecionado[0], "values")
        id_registro = valores[0]

        self.txt_preview.config(state="normal")
        self.txt_preview.delete("1.0", tk.END)
        self.txt_historico.config(state="normal")
        self.txt_historico.delete("1.0", tk.END)

        conn = conectar_db()
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM gestantes WHERE id=?", (id_registro,))
        linha = cursor.fetchone()
        conn.close()

        if linha:
            texto = f"🤰 ACOMPANHAMENTO DE GESTANTE\n\n"
            texto += f"   Empresa: {linha['empresa']} (Cód: {linha['cod_empresa'] or '-'})\n"
            texto += f"   Funcionária: {linha['nome_empregado']} (Cód: {linha['cod_empregado'] or '-'})\n"
            texto += f"   Status Atual: {linha['status']}\n\n"
            
            texto += f"📅 DADOS DO EXAME E PROJEÇÕES\n"
            texto += f"   Data do Exame Constatado: {linha['data_exame']}\n"
            texto += f"   Semanas Constatadas: {linha['semanas_exame']}\n"
            texto += f"   Projeção 40ª Semana (Parto): {linha['data_37_semanas']}\n"
            texto += f"   Fim da Estabilidade: {linha['estabilidade']}\n"

            self.txt_preview.insert("1.0", texto)

        self.txt_preview.config(state="disabled")
        self.txt_historico.config(state="disabled")

    def ordenar_coluna(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        if col in ("Últ. Dia", "Obrigação"):
            def formatar_data_sort(x):
                try: return datetime.strptime(x[0], "%d/%m/%Y")
                except ValueError: return datetime.min if not reverse else datetime.max
            l.sort(key=formatar_data_sort, reverse=reverse)
        elif col == "ID": l.sort(key=lambda t: int(t[0]), reverse=reverse)
        else: l.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
        for index, (val, k) in enumerate(l): self.tree.move(k, '', index)
        self.tree.heading(col, command=lambda _col=col: self.ordenar_coluna(_col, not reverse))

    def abrir_cadastro(self): TelaCadastro(self.root, self.carregar_dados)
    def chamar_importacao(self): executar_importacao(self.root, self.carregar_dados)
    
    def abrir_manutencao(self, event):
        selec = self.tree.selection()
        if not selec: return
        TelaManutencao(self.root, self.tree.item(selec[0], "values")[0], self.carregar_dados)

    def abrir_manutencao_gestante(self, event):
        selec = self.tree_gest.selection()
        if not selec: return
        TelaManutencaoGestante(self.root, self.tree_gest.item(selec[0], "values")[0], self.carregar_dados)

if __name__ == "__main__":
    if not os.path.exists('afastamentos.db'): inicializar_db()
    else: inicializar_db()
    root = tk.Tk()
    app = SistemaAfastamentos(root)
    root.mainloop()