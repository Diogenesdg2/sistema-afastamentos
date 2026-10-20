import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime, timedelta
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

class TelaGestante:
    def __init__(self, parent, callback_atualizar):
        self.callback_atualizar = callback_atualizar
        
        self.janela = tk.Toplevel(parent)
        self.janela.title("Registar Nova Gestante (Acompanhamento)")
        self.janela.geometry("600x480") 
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
        self.ent_empresa = tk.Entry(box_dados, width=28, font=fonte_entry, relief="solid", bd=1)
        self.ent_empresa.grid(row=0, column=1, padx=(5, 15), sticky="w", pady=8)

        tk.Label(box_dados, text="Cód. Empresa:", bg="white", font=fonte_label).grid(row=0, column=2, sticky="e", pady=8)
        self.ent_cod_empresa = tk.Entry(box_dados, width=10, font=fonte_entry, relief="solid", bd=1)
        self.ent_cod_empresa.grid(row=0, column=3, sticky="w", padx=5, pady=8)

        tk.Label(box_dados, text="Funcionária:", bg="white", font=fonte_label).grid(row=1, column=0, sticky="w", pady=8)
        self.ent_nome = tk.Entry(box_dados, width=28, font=fonte_entry, relief="solid", bd=1)
        self.ent_nome.grid(row=1, column=1, padx=(5, 15), sticky="w", pady=8)

        tk.Label(box_dados, text="Cód. Func.:", bg="white", font=fonte_label).grid(row=1, column=2, sticky="e", pady=8)
        self.ent_cod_func = tk.Entry(box_dados, width=10, font=fonte_entry, relief="solid", bd=1)
        self.ent_cod_func.grid(row=1, column=3, sticky="w", padx=5, pady=8)

        ttk.Separator(box_dados, orient='horizontal').grid(row=2, column=0, columnspan=4, sticky="ew", pady=15)

        tk.Label(box_dados, text="Data do Exame:", bg="white", font=fonte_label).grid(row=3, column=0, sticky="w", pady=8)
        self.ent_data_exame = DateEntryLivre(box_dados, width=15, background='#DB2777', foreground='white', borderwidth=0, date_pattern='dd/mm/yyyy', state="normal", font=fonte_entry)
        self.ent_data_exame.grid(row=3, column=1, padx=(5, 15), sticky="w", pady=8)
        self.ent_data_exame.delete(0, tk.END)
        self.ent_data_exame.bind("<<DateEntrySelected>>", self.calcular_projecoes)
        self.ent_data_exame.bind("<KeyRelease>", self.calcular_projecoes)

        tk.Label(box_dados, text="Semanas Constatadas:", bg="white", font=fonte_label).grid(row=3, column=2, sticky="e", pady=8)
        self.ent_semanas = tk.Entry(box_dados, width=10, font=fonte_entry, relief="solid", bd=1, justify="center")
        self.ent_semanas.grid(row=3, column=3, sticky="w", padx=5, pady=8)
        self.ent_semanas.bind("<KeyRelease>", self.calcular_projecoes)

        f_proj = tk.Frame(box_dados, bg="#FDF2F8", bd=1, relief="solid", padx=10, pady=10)
        f_proj.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(15, 0))

        tk.Label(f_proj, text="Projeção 40ª Semana (Parto):", bg="#FDF2F8", font=fonte_label, fg="#BE185D").grid(row=0, column=0, sticky="w")
        self.lbl_40_sem = tk.Label(f_proj, text="-", bg="#FDF2F8", font=("Segoe UI", 10, "bold"), fg="#9D174D")
        self.lbl_40_sem.grid(row=0, column=1, sticky="w", padx=10)

        tk.Label(f_proj, text="Estabilidade:", bg="#FDF2F8", font=fonte_label, fg="#BE185D").grid(row=1, column=0, sticky="w", pady=(5,0))
        self.lbl_estab = tk.Label(f_proj, text="-", bg="#FDF2F8", font=("Segoe UI", 10, "bold"), fg="#9D174D")
        self.lbl_estab.grid(row=1, column=1, sticky="w", padx=10, pady=(5,0))

        btn_salvar = tk.Button(self.janela, text="✔ Registar Gestante", command=self.salvar, 
                               bg="#DB2777", fg="white", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", pady=8)
        btn_salvar.pack(fill="x", padx=20, pady=(0, 20))

        self.data_40_calc = ""
        self.data_estab_calc = ""

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

        if not nome or not emp or not dt_exame or not semanas:
            messagebox.showwarning("Aviso", "Preencha Empresa, Nome, Data do Exame e Semanas.", parent=self.janela)
            return

        try:
            conn = conectar_db()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO gestantes (
                    empresa, cod_empresa, cod_empregado, nome_empregado,
                    semanas_exame, data_exame, data_37_semanas, estabilidade, alerta_nasc_tratado, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0, 'Gestante')
            ''', (emp, cod_emp, cod_func, nome, float(semanas), dt_exame, self.data_40_calc, self.data_estab_calc))
            conn.commit()
            conn.close()

            messagebox.showinfo("Sucesso", f"Gestante {nome} em acompanhamento!\nData Provável do Parto (40 Sem): {self.data_40_calc}", parent=self.janela)
            if self.callback_atualizar: self.callback_atualizar()
            self.janela.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao guardar:\n{str(e)}", parent=self.janela)