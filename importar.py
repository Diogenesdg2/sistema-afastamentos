import pandas as pd
from datetime import datetime
from tkinter import filedialog, messagebox
from database import conectar_db

def formatar_data(val):
    if pd.isna(val): return ""
    if isinstance(val, datetime): return val.strftime('%d/%m/%Y')
    return str(val).split()[0] if str(val).strip() else ""

def executar_importacao(janela_pai, callback_atualizar):
    arquivo = filedialog.askopenfilename(
        parent=janela_pai, 
        title="Selecione a folha de cálculo", 
        filetypes=[("Excel", "*.xlsx *.xls")]
    )
    if not arquivo: return

    try:
        df = pd.read_excel(arquivo, sheet_name=0, header=1) 
        conn = conectar_db()
        cursor = conn.cursor()
        importados = 0
        
        for index, row in df.iterrows():
            def valor_seguro(idx): return row.iloc[idx] if idx < len(row) else None

            nome_val = valor_seguro(3)
            nome = str(nome_val) if pd.notna(nome_val) else ""
            if not nome or nome.lower() == 'nan': continue
            
            empresa = str(valor_seguro(1)) if pd.notna(valor_seguro(1)) else ""
            cod_emp = str(valor_seguro(2)).replace('.0','') if pd.notna(valor_seguro(2)) else ""
            
            tipo = "Outros"
            if pd.notna(valor_seguro(4)): tipo = "Doença"
            elif pd.notna(valor_seguro(5)): tipo = "Acidente"
            elif pd.notna(valor_seguro(6)): tipo = "Cárcere"
            elif pd.notna(valor_seguro(7)): tipo = "Militar"

            ult_dia = formatar_data(valor_seguro(8))
            qtd_dias = str(valor_seguro(9)).replace('.0','') if pd.notna(valor_seguro(9)) else ""
            per_de = formatar_data(valor_seguro(10))
            per_ate = formatar_data(valor_seguro(11))
            d16 = formatar_data(valor_seguro(12))
            cid = str(valor_seguro(13)) if pd.notna(valor_seguro(13)) else ""
            aposentado = str(valor_seguro(14)) if pd.notna(valor_seguro(14)) else ""
            obrigacao = formatar_data(valor_seguro(15))
            ret_data = formatar_data(valor_seguro(16))
            an_dist = formatar_data(valor_seguro(17))
            atend_inss = formatar_data(valor_seguro(18))
            per_pres = formatar_data(valor_seguro(19))
            conc_ben = str(valor_seguro(20)).replace('.0','') if pd.notna(valor_seguro(20)) else ""
            conc_de = formatar_data(valor_seguro(21))
            conc_ate = formatar_data(valor_seguro(22))
            recurso = formatar_data(valor_seguro(23))
            nova_an_dist = formatar_data(valor_seguro(24))
            novo_atend = formatar_data(valor_seguro(25))
            nova_per = formatar_data(valor_seguro(26))
            novo_ben = str(valor_seguro(27)).replace('.0','') if pd.notna(valor_seguro(27)) else ""
            nova_per_de = formatar_data(valor_seguro(28))
            nova_per_ate = formatar_data(valor_seguro(29))
            estabilidade = formatar_data(valor_seguro(30))
            acid_trab = str(valor_seguro(31)) if pd.notna(valor_seguro(31)) else ""
            email = str(valor_seguro(33)) if pd.notna(valor_seguro(33)) else ""
            retornou = 1 if pd.notna(valor_seguro(35)) or ret_data else 0

            cursor.execute("SELECT id FROM afastamentos WHERE nome_empregado=? AND empresa=?", (nome, empresa))
            if cursor.fetchone() is None:
                cursor.execute('''
                    INSERT INTO afastamentos (
                        empresa, cod_empregado, nome_empregado, tipo_afastamento, ultimo_dia_trab, qtd_dias, 
                        periodo_de, periodo_ate, decimo_sexto_dia, cid, aposentado, data_obrigacao,
                        retorno_data, analise_distancia, atendimento_inss, pericia_presencial,
                        concessao_beneficio, concessao_de, concessao_ate, recurso_revisao,
                        nova_analise_distancia, novo_atendimento_inss, nova_pericia_data, nova_pericia_beneficio,
                        nova_pericia_de, nova_pericia_ate, estabilidade, acidente_trab, email_enviado, retornou
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''', (empresa, cod_emp, nome, tipo, ult_dia, qtd_dias, per_de, per_ate, d16, cid, aposentado, 
                      obrigacao, ret_data, an_dist, atend_inss, per_pres, conc_ben, conc_de, conc_ate, recurso, 
                      nova_an_dist, novo_atend, nova_per, novo_ben, nova_per_de, nova_per_ate, estabilidade, acid_trab, email, retornou))
                importados += 1
                
        conn.commit()
        conn.close()
        
        # Chama a função de carregar_dados do app.py para atualizar o ecrã
        callback_atualizar()
        messagebox.showinfo("Sucesso", f"{importados} registos importados com sucesso.", parent=janela_pai)
        
    except Exception as e:
        messagebox.showerror("Erro", f"Falha na importação:\n{str(e)}", parent=janela_pai)