import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import sqlite3
from database import conectar_db

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

class TelaManutencaoGestante:
    def __init__(self, parent, id_registro, callback_atualizar):
        self.id_registro = id_registro
        self.callback_atualizar = callback_atualizar
        
        self.janela = tk.Toplevel(parent)
        self.janela.withdraw() 
        
        self.dados = self.buscar_dados()
        if not self.dados:
            messagebox.showerror("Erro", "Registo de gestante não encontrado.", parent=self.janela)
            self.janela.destroy()
            return

        nome_func = self.dados['nome_empregado'] or "Sem Nome"
        self.janela.title(f"Manutenção de Gestante | ID: {self.id_registro} - {nome_func}")
        self.janela.geometry("650x550") 
        self.janela.grab_set()
        self.janela.resizable(False, False)
        self.janela.configure(bg="#F3F4F6")

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TLabelframe", background="#FFFFFF", borderwidth=2)
        style.configure("TLabelframe.Label", background="#FFFFFF", font=("Segoe UI", 10, "bold"), foreground="#DB2777")

        fonte_label = ("Segoe UI", 9, "bold")
        fonte_entry = ("Segoe UI", 9)

        box_dados = ttk.LabelFrame(self.janela, text=" 🤰 Dados da Gestante e do Exame ", padding=20)
        box_dados.pack(fill="both", expand=True, padx=20, pady=20)

        tk.Label(box_dados, text="Empresa:", bg="white", font=fonte_label).grid(row=0, column=0, sticky="w", pady=8)
        self.ent_empresa = tk.Entry(box_dados, width=32, font=fonte_entry, relief="solid", bd=1)
        self.ent_empresa.grid(row=0, column=1, padx=(5, 15), sticky="w", pady=8)

        tk.Label(box_dados, text="Cód. Empresa:", bg="white", font=fonte_label).grid(row=0, column=2, sticky="e", pady=8)
        self.ent_cod_empresa = tk.Entry(box_dados, width=12, font=fonte_entry, relief="solid", bd=1)
        self.ent_cod_empresa.grid(row=0, column=3, sticky="w", padx=5, pady=8)

        tk.Label(box_dados, text="Funcionária:", bg="white", font=fonte_label).grid(row=1, column=0, sticky="w", pady=8)
        self.ent_nome = tk.Entry(box_dados, width=32, font=fonte_entry, relief="solid", bd=1)
        self.ent_nome.grid(row=1, column=1, padx=(5, 15), sticky="w", pady=8)

        tk.Label(box_dados, text="Cód. Func.:", bg="white", font=fonte_label).grid(row=1, column=2, sticky="e", pady=8)
        self.ent_cod_func = tk.Entry(box_dados, width=12, font=fonte_entry, relief="solid", bd=1)
        self.ent_cod_func.grid(row=1, column=3, sticky="w", padx=5, pady=8)

        tk.Label(box_dados, text="Status Atual:", bg="white", font=fonte_label).grid(row=2, column=0, sticky="w", pady=8)
        self.combo_status = ttk.Combobox(box_dados, values=["Gestante", "Afastada", "Concluído"], width=20, state="readonly", font=fonte_entry)
        self.combo_status.grid(row=2, column=1, padx=(5, 15), sticky="w", pady=8)

        ttk.Separator(box_dados, orient='horizontal').grid(row=3, column=0, columnspan=4, sticky="ew", pady=15)

        tk.Label(box_dados, text="Data do Exame:", bg="white", font=fonte_label).grid(row=4, column=0, sticky="w", pady=8)
        self.ent_data_exame = DateEntryLivre(box_dados, width=15, background='#DB2777', foreground='white', borderwidth=0, date_pattern='dd/mm/yyyy', state="normal", font=fonte_entry)
        self.ent_data_exame.grid(row=4, column=1, padx=(5, 15), sticky="w", pady=8)
        self.ent_data_exame.bind("<<DateEntrySelected>>", self.calcular_projecoes)
        self.ent_data_exame.bind("<KeyRelease>", self.calcular_projecoes)

        tk.Label(box_dados, text="Semanas (Exame):", bg="white", font=fonte_label).grid(row=4, column=2, sticky="e", pady=8)
        self.ent_semanas = tk.Entry(box_dados, width=10, font=fonte_entry, relief="solid", bd=1, justify="center")
        self.ent_semanas.grid(row=4, column=3, sticky="w", padx=5, pady=8)
        self.ent_semanas.bind("<KeyRelease>", self.calcular_projecoes)

        f_proj = tk.Frame(box_dados, bg="#FDF2F8", bd=1, relief="solid", padx=10, pady=10)
        f_proj.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(15, 0))

        tk.Label(f_proj, text="Projeção 40ª Semana (Parto):", bg="#FDF2F8", font=fonte_label, fg="#BE185D").grid(row=0, column=0, sticky="w")
        self.lbl_40_sem = tk.Label(f_proj, text="-", bg="#FDF2F8", font=("Segoe UI", 10, "bold"), fg="#9D174D")
        self.lbl_40_sem.grid(row=0, column=1, sticky="w", padx=10)

        tk.Label(f_proj, text="Estabilidade:", bg="#FDF2F8", font=fonte_label, fg="#BE185D").grid(row=1, column=0, sticky="w", pady=(5,0))
        self.lbl_estab = tk.Label(f_proj, text="-", bg="#FDF2F8", font=("Segoe UI", 10, "bold"), fg="#9D174D")
        self.lbl_estab.grid(row=1, column=1, sticky="w", padx=10, pady=(5,0))

        self.var_alerta_nasc = tk.IntVar()
        chk_nasc = tk.Checkbutton(box_dados, text="✔ Marcar Lembrete de Nascimento como Tratado", variable=self.var_alerta_nasc, bg="white", font=fonte_label, fg="#DB2777")
        chk_nasc.grid(row=6, column=0, columnspan=4, sticky="w", pady=(15, 0))

        btn_salvar = tk.Button(self.janela, text="💾 Salvar Alterações", command=self.salvar, 
                               bg="#DB2777", fg="white", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", pady=8)
        btn_salvar.pack(fill="x", padx=20, pady=(0, 20))

        self.data_40_calc = ""
        self.data_estab_calc = ""

        self.preencher_dados()
        self.janela.deiconify()

    def buscar_dados(self):
        conn = conectar_db()
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM gestantes WHERE id=?", (self.id_registro,))
        linha = cursor.fetchone()
        conn.close()
        return linha

    def preencher_dados(self):
        self.ent_empresa.insert(0, self.dados['empresa'] or "")
        self.ent_cod_empresa.insert(0, self.dados['cod_empresa'] or "")
        self.ent_nome.insert(0, self.dados['nome_empregado'] or "")
        self.ent_cod_func.insert(0, self.dados['cod_empregado'] or "")
        self.combo_status.set(self.dados['status'] or "Gestante")
        self.ent_semanas.insert(0, str(self.dados['semanas_exame']) if self.dados['semanas_exame'] else "")
        self.var_alerta_nasc.set(self.dados['alerta_nasc_tratado'] or 0)

        dt_ex = self.dados['data_exame']
        if dt_ex:
            self.ent_data_exame.delete(0, tk.END)
            self.ent_data_exame.insert(0, dt_ex)
            self.calcular_projecoes()

    def calcular_projecoes(self, event=None):
        dt_exame_str = self.ent_data_exame.get().strip()
        semanas_str = self.ent_semanas.get().strip().replace(",", ".")

        if dt_exame_str and semanas_str:
            try:
                dt_exame = datetime.strptime(dt_exame_str, "%d/%m/%Y")
                semanas = float(semanas_str)
                dias_para_40 = (40.0 - semanas) * 7
                dt_40 = dt_exame + timedelta(days=dias_para_40)
                self.data_40_calc = dt_40.strftime("%d/%m/%Y")
                self.lbl_40_sem.config(text=self.data_40_calc)

                dt_estab = dt_exame + timedelta(days=365)
                self.data_estab_calc = dt_estab.strftime("%d/%m/%Y")
                self.lbl_estab.config(text=self.data_estab_calc)
            except ValueError:
                self.lbl_40_sem.config(text="-")
                self.lbl_estab.config(text="-")
                self.data_40_calc = ""
                self.data_estab_calc = ""
        else:
            self.lbl_40_sem.config(text="-")
            self.lbl_estab.config(text="-")
            self.data_40_calc = ""
            self.data_estab_calc = ""

    def salvar(self):
        emp = self.ent_empresa.get().strip()
        cod_emp = self.ent_cod_empresa.get().strip()
        nome = self.ent_nome.get().strip()
        cod_func = self.ent_cod_func.get().strip()
        dt_exame = self.ent_data_exame.get().strip()
        semanas = self.ent_semanas.get().strip().replace(",", ".")
        status = self.combo_status.get()
        tratado = self.var_alerta_nasc.get()

        if not nome or not emp or not dt_exame or not semanas:
            messagebox.showwarning("Aviso", "Preencha Empresa, Nome, Data do Exame e Semanas.", parent=self.janela)
            return

        # Recolhe os dados originais antes da edição para encontrar o afastamento associado
        old_nome = self.dados['nome_empregado']
        old_empresa = self.dados['empresa']

        try:
            conn = conectar_db()
            cursor = conn.cursor()
            
            # 1. Atualiza a tabela Gestantes
            cursor.execute('''
                UPDATE gestantes SET
                    empresa=?, cod_empresa=?, cod_empregado=?, nome_empregado=?,
                    semanas_exame=?, data_exame=?, data_37_semanas=?, estabilidade=?, 
                    alerta_nasc_tratado=?, status=?
                WHERE id=?
            ''', (emp, cod_emp, cod_func, nome, float(semanas), dt_exame, self.data_40_calc, self.data_estab_calc, tratado, status, self.id_registro))
            
            # 2. Sincroniza automaticamente com a tabela Afastamentos
            cursor.execute('''
                UPDATE afastamentos SET
                    empresa=?, cod_empresa=?, cod_empregado=?, nome_empregado=?
                WHERE nome_empregado=? AND empresa=? AND tipo_afastamento='Maternidade'
            ''', (emp, cod_emp, cod_func, nome, old_nome, old_empresa))
            
            conn.commit()
            conn.close()

            messagebox.showinfo("Sucesso", "Dados da gestante atualizados com sucesso!", parent=self.janela)
            if self.callback_atualizar: self.callback_atualizar()
            self.janela.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar:\n{str(e)}", parent=self.janela)