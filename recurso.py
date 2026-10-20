import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from database import conectar_db

class TelaRecurso:
    def __init__(self, parent, id_afastamento, empresa, funcionario):
        self.id_afastamento = id_afastamento
        self.empresa = empresa
        self.funcionario = funcionario

        self.janela = tk.Toplevel(parent)
        self.janela.title("Cadastrar Recurso / Revisão / Contestação")
        self.janela.geometry("600x420")
        self.janela.grab_set()
        self.janela.configure(bg="#F3F4F6")
        
        self.construir_tela()

    def construir_tela(self):
        lbl_titulo = tk.Label(self.janela, text="Novo Recurso / Revisão / Contestação", font=("Segoe UI", 12, "bold"), bg="#F3F4F6", fg="#D97706")
        lbl_titulo.pack(pady=10)

        box_info = ttk.LabelFrame(self.janela, text=" Informações do Afastamento ", padding=15)
        box_info.pack(fill="x", padx=20, pady=5)
        
        tk.Label(box_info, text=f"ID Afastamento Principal: {self.id_afastamento}", font=("Segoe UI", 10, "bold"), bg="white", fg="#0284C7").grid(row=0, column=0, sticky="w", pady=2)
        tk.Label(box_info, text=f"Empresa: {self.empresa}", font=("Segoe UI", 10), bg="white").grid(row=1, column=0, sticky="w", pady=2)
        tk.Label(box_info, text=f"Funcionário: {self.funcionario}", font=("Segoe UI", 10), bg="white").grid(row=2, column=0, sticky="w", pady=2)

        box_vinc = ttk.LabelFrame(self.janela, text=" Vincular à Origem (Análise ou Perícia) ", padding=15)
        box_vinc.pack(fill="x", padx=20, pady=10)

        tk.Label(box_vinc, text="Tipo de Vínculo:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.combo_tipo = ttk.Combobox(box_vinc, values=["Análise Documental", "Perícia Presencial"], state="readonly", width=25)
        self.combo_tipo.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        tk.Label(box_vinc, text="Item Específico:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        self.combo_item = ttk.Combobox(box_vinc, state="readonly", width=45)
        self.combo_item.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        
        self.mapa_itens = {}

        self.combo_tipo.bind("<<ComboboxSelected>>", self.carregar_itens)

        tk.Button(self.janela, text="💾 Cadastrar Recurso Inicial", command=self.salvar_recurso, bg="#10B981", fg="white", font=("Segoe UI", 10, "bold"), padx=20, pady=5, relief="flat", cursor="hand2").pack(pady=15)

    def carregar_itens(self, event=None):
        self.combo_item.set("")
        self.mapa_itens.clear()
        tipo = self.combo_tipo.get()
        
        conn = conectar_db()
        cursor = conn.cursor()
        
        if tipo == "Análise Documental":
            cursor.execute("SELECT id, data_analise, numero_beneficio, situacao FROM analises_inss WHERE id_afastamento=?", (self.id_afastamento,))
            for r in cursor.fetchall():
                lbl = f"ID: {r[0]} | Data: {r[1]} | Ben: {r[2]} | {r[3]}"
                self.mapa_itens[lbl] = r[0]
        elif tipo == "Perícia Presencial":
            cursor.execute("SELECT id, data_atendimento, numero_beneficio, situacao FROM pericias_inss WHERE id_afastamento=?", (self.id_afastamento,))
            for r in cursor.fetchall():
                lbl = f"ID: {r[0]} | Atend: {r[1]} | Ben: {r[2]} | {r[3]}"
                self.mapa_itens[lbl] = r[0]
        
        conn.close()
        self.combo_item['values'] = list(self.mapa_itens.keys())
        if self.mapa_itens:
            self.combo_item.current(0)

    def salvar_recurso(self):
        tipo = self.combo_tipo.get()
        item_lbl = self.combo_item.get()
        if not tipo or not item_lbl:
            messagebox.showwarning("Aviso", "Selecione o Tipo de Vínculo e o Item Específico na lista.", parent=self.janela)
            return
        
        id_vinc = self.mapa_itens.get(item_lbl)
        dt_hoje = datetime.now().strftime("%d/%m/%Y")
        
        try:
            conn = conectar_db()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO recursos_inss (id_afastamento, tipo_vinculo, id_vinculo, data_criacao, status)
                VALUES (?, ?, ?, ?, ?)
            ''', (self.id_afastamento, tipo, id_vinc, dt_hoje, 'Aberto'))
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Sucesso", "Recurso cadastrado com sucesso! Ele aparecerá na nova aba da tela principal.", parent=self.janela)
            self.janela.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao cadastrar recurso: {str(e)}", parent=self.janela)