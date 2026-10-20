import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tkcalendar import DateEntry
from datetime import datetime
import sqlite3
import base64
from database import conectar_db

CHAVE_MESTRE = "VimanSeguranca"

def desencriptar_senha(encoded):
    if not encoded: return ""
    if not encoded.startswith("ENC::"): return encoded 
    try:
        encoded_str = encoded[5:]
        xored = base64.b64decode(encoded_str.encode('utf-8')).decode('utf-8')
        return ''.join(chr(ord(c) ^ ord(k)) for c, k in zip(xored, CHAVE_MESTRE * (len(xored)//len(CHAVE_MESTRE) + 1)))
    except:
        return encoded

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

class TelaManutencaoRevisao:
    def __init__(self, parent, id_recurso, callback_atualizar):
        self.id_recurso = id_recurso
        self.callback_atualizar = callback_atualizar
        
        self.janela = tk.Toplevel(parent)
        self.janela.title(f"Gestão de Revisão / Recurso | ID: {self.id_recurso}")
        self.janela.geometry("850x700")
        self.janela.grab_set()
        self.janela.configure(bg="#F3F4F6")
        
        self.dados = self.buscar_dados_completos()
        if not self.dados:
            messagebox.showerror("Erro", "Recurso não encontrado no banco de dados.", parent=self.janela)
            self.janela.destroy()
            return

        self.construir_tela()
        self.preencher_dados()

    def buscar_dados_completos(self):
        conn = conectar_db()
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # Busca os dados do recurso cruzando com o afastamento (agora puxando tbm os dados do gov.br)
        query = """
            SELECT r.*, a.empresa, a.nome_empregado, a.tipo_afastamento, a.id as id_afast,
                   a.gov_br_autorizacao, a.cpf, a.gov_br_senha
            FROM recursos_inss r
            JOIN afastamentos a ON r.id_afastamento = a.id
            WHERE r.id = ?
        """
        cursor.execute(query, (self.id_recurso,))
        linha = cursor.fetchone()
        
        if not linha:
            conn.close()
            return None
            
        dados_dit = dict(linha)
        
        # Busca CIDs relacionados
        cursor.execute("SELECT codigo FROM cids_historico WHERE id_afastamento=?", (dados_dit['id_afast'],))
        cids = cursor.fetchall()
        dados_dit['str_cids'] = " | ".join([c[0] for c in cids]) if cids else "Nenhum"
        
        # Busca o número do Benefício relacionado à origem do recurso
        dados_dit['numero_beneficio'] = "Não Localizado"
        if dados_dit['tipo_vinculo'] == 'Análise Documental':
            cursor.execute("SELECT numero_beneficio FROM analises_inss WHERE id=?", (dados_dit['id_vinculo'],))
            ben = cursor.fetchone()
            if ben and ben[0]: dados_dit['numero_beneficio'] = ben[0]
        elif dados_dit['tipo_vinculo'] == 'Perícia Presencial':
            cursor.execute("SELECT numero_beneficio FROM pericias_inss WHERE id=?", (dados_dit['id_vinculo'],))
            ben = cursor.fetchone()
            if ben and ben[0]: dados_dit['numero_beneficio'] = ben[0]
            
        conn.close()
        return dados_dit

    def construir_tela(self):
        fonte_lbl = ("Segoe UI", 10, "bold")
        fonte_ent = ("Segoe UI", 10)
        
        # --- BLOCO 1: DADOS INTOCÁVEIS (Apenas Leitura) ---
        box_readonly = ttk.LabelFrame(self.janela, text=" Informações de Origem (Apenas Leitura) ", padding=15)
        box_readonly.pack(fill="x", padx=20, pady=(15, 10))
        
        tk.Label(box_readonly, text="Empresa:", bg="white", font=fonte_lbl).grid(row=0, column=0, sticky="w", pady=5)
        self.ent_emp = tk.Entry(box_readonly, width=35, font=fonte_ent, relief="solid", bd=1, bg="#E5E7EB")
        self.ent_emp.grid(row=0, column=1, sticky="w", padx=(5, 20), pady=5)
        
        tk.Label(box_readonly, text="Funcionário:", bg="white", font=fonte_lbl).grid(row=0, column=2, sticky="w", pady=5)
        self.ent_func = tk.Entry(box_readonly, width=35, font=fonte_ent, relief="solid", bd=1, bg="#E5E7EB")
        self.ent_func.grid(row=0, column=3, sticky="w", padx=5, pady=5)
        
        tk.Label(box_readonly, text="Motivo Afast.:", bg="white", font=fonte_lbl).grid(row=1, column=0, sticky="w", pady=5)
        self.ent_motivo = tk.Entry(box_readonly, width=20, font=fonte_ent, relief="solid", bd=1, bg="#E5E7EB")
        self.ent_motivo.grid(row=1, column=1, sticky="w", padx=(5, 20), pady=5)
        
        tk.Label(box_readonly, text="CIDs Relacionados:", bg="white", font=fonte_lbl).grid(row=1, column=2, sticky="w", pady=5)
        self.ent_cids = tk.Entry(box_readonly, width=35, font=fonte_ent, relief="solid", bd=1, bg="#E5E7EB")
        self.ent_cids.grid(row=1, column=3, sticky="w", padx=5, pady=5)

        tk.Label(box_readonly, text="Origem Vinculada:", bg="white", font=fonte_lbl).grid(row=2, column=0, sticky="w", pady=5)
        self.ent_origem = tk.Entry(box_readonly, width=35, font=fonte_ent, relief="solid", bd=1, bg="#E0F2FE", fg="#0284C7")
        self.ent_origem.grid(row=2, column=1, sticky="w", padx=(5, 20), pady=5)
        
        tk.Label(box_readonly, text="Nº Benefício:", bg="white", font=fonte_lbl).grid(row=2, column=2, sticky="w", pady=5)
        self.ent_ben = tk.Entry(box_readonly, width=25, font=("Segoe UI", 10, "bold"), relief="solid", bd=1, bg="#E0F2FE", fg="#0284C7")
        self.ent_ben.grid(row=2, column=3, sticky="w", padx=5, pady=5)

        # --- BLOCO 1.5: ACESSO GOV.BR ---
        box_gov = ttk.LabelFrame(self.janela, text=" Autorização de Acesso (GOV.BR) ", padding=10)
        box_gov.pack(fill="x", padx=20, pady=(0, 15))

        tk.Label(box_gov, text="Tem Autorização no GOV.BR?:", bg="white", font=fonte_lbl).pack(side="left", padx=5)
        self.ent_gov_auth = tk.Entry(box_gov, width=10, font=fonte_ent, relief="solid", bd=1, bg="#E5E7EB")
        self.ent_gov_auth.pack(side="left", padx=5)
        
        tk.Label(box_gov, text="CPF:", bg="white", font=fonte_lbl).pack(side="left", padx=(20, 5))
        self.ent_gov_cpf = tk.Entry(box_gov, width=16, font=fonte_ent, relief="solid", bd=1, bg="#E5E7EB")
        self.ent_gov_cpf.pack(side="left", padx=5)

        tk.Label(box_gov, text="Senha de Acesso:", bg="white", font=fonte_lbl).pack(side="left", padx=(20, 5))
        self.ent_gov_senha = tk.Entry(box_gov, width=20, font=fonte_ent, relief="solid", bd=1, bg="#E5E7EB", show="*")
        self.ent_gov_senha.pack(side="left", padx=5)

        tk.Button(box_gov, text="👁️", command=self.pedir_senha_mestre, bg="white", relief="flat", cursor="hand2").pack(side="left")

        # --- BLOCO 2: DADOS EDITÁVEIS DO RECURSO ---
        box_edit = ttk.LabelFrame(self.janela, text=" Gestão do Recurso / Contestação ", padding=15)
        box_edit.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        
        f_linha1 = tk.Frame(box_edit, bg="white")
        f_linha1.pack(fill="x", pady=5)
        
        tk.Label(f_linha1, text="Criado em:", bg="white", font=fonte_lbl, fg="#6B7280").pack(side="left")
        self.lbl_criado = tk.Label(f_linha1, text="", bg="white", font=("Segoe UI", 10, "bold"), fg="#6B7280")
        self.lbl_criado.pack(side="left", padx=(5, 30))
        
        tk.Label(f_linha1, text="Última Consulta:", bg="white", font=fonte_lbl).pack(side="left")
        self.ent_dt_consulta = DateEntryLivre(f_linha1, width=12, background='#F59E0B', foreground='white', borderwidth=0, date_pattern='dd/mm/yyyy', font=fonte_ent)
        self.ent_dt_consulta.pack(side="left", padx=(5, 30))

        tk.Label(f_linha1, text="Tipo:", bg="white", font=fonte_lbl).pack(side="left")
        self.combo_tipo = ttk.Combobox(f_linha1, values=["Revisão", "Recurso", "Contestação", "Ação Judicial"], state="readonly", width=15, font=fonte_ent)
        self.combo_tipo.pack(side="left", padx=(5, 30))

        tk.Label(f_linha1, text="Status:", bg="white", font=fonte_lbl).pack(side="left")
        self.combo_status = ttk.Combobox(f_linha1, values=["Aberto", "Deferido", "Indeferido"], state="readonly", width=12, font=fonte_ent)
        self.combo_status.pack(side="left", padx=5)

        tk.Label(box_edit, text="Observações e Andamentos do Processo:", bg="white", font=fonte_lbl).pack(anchor="w", pady=(15, 5))
        self.txt_obs = tk.Text(box_edit, width=80, height=12, font=fonte_ent, relief="solid", bd=1)
        self.txt_obs.pack(fill="both", expand=True)

        # --- BOTÃO SALVAR ---
        f_botoes = tk.Frame(self.janela, bg="#F3F4F6")
        f_botoes.pack(fill="x", padx=20, pady=(0, 20))
        
        tk.Button(f_botoes, text="💾 Salvar Alterações do Recurso", command=self.salvar_alteracoes, 
                  bg="#10B981", fg="white", font=("Segoe UI", 11, "bold"), padx=40, pady=8, relief="flat", cursor="hand2").pack()

    def pedir_senha_mestre(self):
        if not self.ent_gov_senha.get(): return
        if self.ent_gov_senha.cget('show') == '':
            self.ent_gov_senha.config(show="*")
            return
        pwd = simpledialog.askstring("Segurança", "Digite a senha mestre para visualizar:", parent=self.janela, show='*')
        if pwd == "@Viman1234@":
            self.ent_gov_senha.config(show="")
        elif pwd is not None:
            messagebox.showerror("Acesso Negado", "Senha mestre incorreta.", parent=self.janela)

    def preencher_dados(self):
        # 1. Preencher Read-only do Afastamento
        for widget, valor in [
            (self.ent_emp, self.dados['empresa']),
            (self.ent_func, self.dados['nome_empregado']),
            (self.ent_motivo, self.dados['tipo_afastamento']),
            (self.ent_cids, self.dados['str_cids']),
            (self.ent_origem, f"{self.dados['tipo_vinculo']} (ID: {self.dados['id_vinculo']})"),
            (self.ent_ben, self.dados['numero_beneficio'])
        ]:
            widget.insert(0, valor or "")
            widget.config(state="readonly")
            
        # 1.5 Preencher Read-only do GOV.BR
        self.ent_gov_auth.insert(0, self.dados['gov_br_autorizacao'] or "")
        self.ent_gov_auth.config(state="readonly")
        
        self.ent_gov_cpf.insert(0, self.dados['cpf'] or "")
        self.ent_gov_cpf.config(state="readonly")
        
        senha_descriptografada = desencriptar_senha(self.dados['gov_br_senha'] or "")
        self.ent_gov_senha.insert(0, senha_descriptografada)
        self.ent_gov_senha.config(state="readonly")

        # 2. Preencher Editáveis
        self.lbl_criado.config(text=self.dados['data_criacao'])
        
        if self.dados['ultima_consulta']:
            try:
                self.ent_dt_consulta.set_date(datetime.strptime(self.dados['ultima_consulta'], "%d/%m/%Y"))
            except Exception: pass
        else:
            self.ent_dt_consulta.delete(0, tk.END)

        self.combo_tipo.set(self.dados['tipo_recurso'] or "")
        self.combo_status.set(self.dados['status'] or "Aberto")
        
        self.txt_obs.insert("1.0", self.dados['observacoes'] or "")

    def salvar_alteracoes(self):
        dt_consulta = self.ent_dt_consulta.get()
        tipo = self.combo_tipo.get()
        status = self.combo_status.get()
        obs = self.txt_obs.get("1.0", tk.END).strip()
        
        try:
            conn = conectar_db()
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE recursos_inss 
                SET ultima_consulta=?, tipo_recurso=?, status=?, observacoes=?
                WHERE id=?
            ''', (dt_consulta, tipo, status, obs, self.id_recurso))
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Sucesso", "Recurso atualizado com sucesso!", parent=self.janela)
            self.callback_atualizar()
            self.janela.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar o recurso: {str(e)}", parent=self.janela)