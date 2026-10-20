import os
import platform
import subprocess
import calendar
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox, filedialog

from reportlab.graphics.shapes import Drawing, Path
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib import colors

from database import conectar_db, obter_ids_mesmo_funcionario


# ================= FUNÇÕES AUXILIARES DE CÁLCULO E DESENHO =================

def data_por_extenso(data_obj):
    """Converte um objeto date para o formato por extenso (ex: 21 de novembro de 2023)."""
    meses = [
        "", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
    dia = data_obj.day
    mes = meses[data_obj.month]
    ano = data_obj.year
    return f"{dia} de {mes} de {ano}"


def create_brace(width):
    """Desenha as chaves horizontais { no quadro do PDF."""
    d = Drawing(width, 22)
    p = Path(fillColor=None, strokeColor=colors.black, strokeWidth=0.8)
    p.moveTo(0, 20)
    p.curveTo(0, 16, 0, 14, 5, 14)
    p.lineTo(width / 2 - 5, 14)
    p.curveTo(width / 2 - 2, 14, width / 2, 12, width / 2, 4)
    p.curveTo(width / 2, 12, width / 2 + 2, 14, width / 2 + 5, 14)
    p.lineTo(width - 5, 14)
    p.curveTo(width, 14, width, 16, width, 20)
    d.add(p)
    return d


def somar_meses(dt, meses):
    """Soma meses exatos a uma data respeitando o último dia de cada mês."""
    mes_total = dt.month - 1 + meses
    ano = dt.year + mes_total // 12
    mes = mes_total % 12 + 1
    ultimo_dia_mes = calendar.monthrange(ano, mes)[1]
    dia = min(dt.day, ultimo_dia_mes)
    return dt.replace(year=ano, month=mes, day=dia)


def calcular_meses_dias_ciclo(dt_ini, dt_fim):
    """Calcula ciclos mensais fechados (ex: 04/11 a 03/12 = 1 mês) e dias restantes."""
    if dt_fim < dt_ini:
        return 0, 0
    meses = 0
    while True:
        prox_inicio = somar_meses(dt_ini, meses + 1)
        fim_ciclo = prox_inicio - timedelta(days=1)
        if fim_ciclo <= dt_fim:
            meses += 1
        else:
            break
    inicio_resto = somar_meses(dt_ini, meses)
    dias_resto = (dt_fim - inicio_resto).days + 1
    return meses, dias_resto


def abrir_pdf_sistema(arquivo):
    """Abre o PDF gerado no visualizador padrão do sistema operacional."""
    if platform.system() == 'Windows':
        os.startfile(arquivo)
    elif platform.system() == 'Darwin':
        subprocess.call(('open', arquivo))
    else:
        subprocess.call(('xdg-open', arquivo))


# ================= ROTEADOR PRINCIPAL =================

def emitir_termo_ferias(tela):
    motivo = tela.combo_motivo.get().strip()
    
    var_ap = getattr(tela, 'var_aposentado', None)
    is_aposentado = False
    if var_ap is not None:
        try:
            is_aposentado = str(var_ap.get()) in ("1", "Sim", "True", "PY_VARIABLE:1")
        except:
            is_aposentado = False

    if is_aposentado and motivo in ["Doença", "Acidente", "Outros"]:
        gerar_termo_aposentado(tela)
    elif motivo in ["Doença", "Acidente", "Outros"]:
        gerar_termo_doenca_acidente(tela)
    elif motivo == "Cárcere":
        gerar_termo_carcere(tela)
    elif motivo == "Militar":
        gerar_termo_militar(tela)
    else:
        messagebox.showwarning("Aviso", f"Não há modelo de termo de férias configurado para o motivo '{motivo}'.", parent=tela.janela)


# ================= CASO 1: DOENÇA E ACIDENTE =================

def gerar_termo_doenca_acidente(tela):
    ret_str = tela.ent_ret_data.get().strip()
    if tela.dias_beneficio_aquisitivo < 180:
        return messagebox.showwarning("Aviso", "O colaborador não acumulou 180 dias de benefício neste período aquisitivo.", parent=tela.janela)
    if not ret_str or len(ret_str) != 10:
        return messagebox.showwarning("Aviso", "Preencha a Data Efetiva de Retorno para emitir o termo.", parent=tela.janela)

    empresa = tela.ent_empresa.get().strip()
    nome = tela.ent_nome.get().strip()
    aq_de = tela.ent_aq_de.get().strip()
    aq_ate = tela.ent_aq_ate.get().strip()
    novo_aq_de = tela.ent_novo_aq_de.get().strip()
    novo_aq_ate = tela.ent_novo_aq_ate.get().strip()
    
    periodo_de_str = tela.ent_per_de.get().strip()
    d16_str = tela.ent_16dia.get().strip()

    if not periodo_de_str and not d16_str:
        return messagebox.showwarning("Aviso", "Para gerar o termo com as datas precisas, certifique-se de que o 'Período De' (início do atestado) ou o '16º Dia' estão preenchidos.", parent=tela.janela)

    conn = conectar_db()
    cursor = conn.cursor()
    ids = obter_ids_mesmo_funcionario(cursor, tela.ent_cpf.get().strip(), tela.ent_cod_empresa.get().strip(), tela.ent_cod_func.get().strip(), tela.id_registro)
    if not ids:
        conn.close()
        return
        
    qmarks = ",".join("?" for _ in ids)
    beneficios = set()
    primeiro_dia_inss = None
    ultimo_dia_inss = None
    
    cursor.execute(f"SELECT numero_beneficio, concessao_de, concessao_ate FROM analises_inss WHERE id_afastamento IN ({qmarks}) AND situacao IN ('Concedido', 'Concedida')", ids)
    for r in cursor.fetchall():
        if r[0]: beneficios.add(r[0])
        if r[1]:
            try:
                dt_b_de = datetime.strptime(r[1], "%d/%m/%Y").date()
                if not primeiro_dia_inss or dt_b_de < primeiro_dia_inss: primeiro_dia_inss = dt_b_de
            except ValueError: pass
        if r[2]:
            try:
                dt_b_ate = datetime.strptime(r[2], "%d/%m/%Y").date()
                if not ultimo_dia_inss or dt_b_ate > ultimo_dia_inss: ultimo_dia_inss = dt_b_ate
            except ValueError: pass

    cursor.execute(f"SELECT numero_beneficio, concessao_de, concessao_ate FROM pericias_inss WHERE id_afastamento IN ({qmarks}) AND situacao IN ('Concedido', 'Concedida')", ids)
    for r in cursor.fetchall():
        if r[0]: beneficios.add(r[0])
        if r[1]:
            try:
                dt_b_de = datetime.strptime(r[1], "%d/%m/%Y").date()
                if not primeiro_dia_inss or dt_b_de < primeiro_dia_inss: primeiro_dia_inss = dt_b_de
            except ValueError: pass
        if r[2]:
            try:
                dt_b_ate = datetime.strptime(r[2], "%d/%m/%Y").date()
                if not ultimo_dia_inss or dt_b_ate > ultimo_dia_inss: ultimo_dia_inss = dt_b_ate
            except ValueError: pass
                
    conn.close()

    try:
        aq_de_dt = datetime.strptime(aq_de, "%d/%m/%Y").date()
        ret_dt = datetime.strptime(ret_str, "%d/%m/%Y").date()
        
        if d16_str and len(d16_str) == 10:
            dt_16_dia = datetime.strptime(d16_str, "%d/%m/%Y").date()
        else:
            dt_per_de = datetime.strptime(periodo_de_str, "%d/%m/%Y").date()
            dt_16_dia = dt_per_de + timedelta(days=15)
            
        dt_15_dia = dt_16_dia - timedelta(days=1)
        quinto_decimo_str = dt_15_dia.strftime("%d/%m/%Y")
        
        inicio_afast_dt = primeiro_dia_inss if primeiro_dia_inss else dt_16_dia
        fim_afast_dt = ultimo_dia_inss if ultimo_dia_inss else (ret_dt - timedelta(days=1))
        
        inicio_afast_str = inicio_afast_dt.strftime("%d/%m/%Y")
        fim_afast_str = fim_afast_dt.strftime("%d/%m/%Y")
        
        meses_trab, dias_trab = calcular_meses_dias_ciclo(aq_de_dt, dt_15_dia)
        if dias_trab >= 15:
            meses_trab += 1
            dias_trab = 0

        if meses_trab > 0 and dias_trab == 0:
            m_str = f"{meses_trab:02d} mês" if meses_trab == 1 else f"{meses_trab:02d} meses"
            trab_label = "trabalhado" if meses_trab == 1 else "trabalhados"
            str_dias_trabalhados = f"{m_str}<br/>{trab_label}"
        elif meses_trab > 0 and dias_trab > 0:
            m_str = f"{meses_trab:02d} mês" if meses_trab == 1 else f"{meses_trab:02d} meses"
            d_str = f" e {dias_trab:02d} dia" if dias_trab == 1 else f" e {dias_trab:02d} dias"
            str_dias_trabalhados = f"{m_str}{d_str}<br/>trabalhados"
        else:
            str_dias_trabalhados = f"{dias_trab:02d} dia<br/>trabalhado" if dias_trab <= 1 else f"{dias_trab:02d} dias<br/>trabalhados"

        meses_afast, dias_afast = calcular_meses_dias_ciclo(inicio_afast_dt, fim_afast_dt)
        if meses_afast > 0:
            m_str = f"{meses_afast:02d} mês" if meses_afast == 1 else f"{meses_afast:02d} meses"
            d_str = f" e {dias_afast:02d} dia" if dias_afast == 1 else (f" e {dias_afast:02d} dias" if dias_afast > 0 else "")
            str_tempo_afast_tbl = f"{m_str}{d_str}<br/>afastamento pelo INSS"
        else:
            str_tempo_afast_tbl = f"{dias_afast:02d} dias<br/>afastamento pelo INSS"
            
    except ValueError:
        return messagebox.showwarning("Erro", "Erro ao calcular datas. Verifique se os campos estão preenchidos num formato válido.", parent=tela.janela)

    str_beneficios = ", ".join([f"nº {b}" for b in beneficios]) if beneficios else "Nº não informado"
    data_hoje_ext = data_por_extenso(datetime.now().date())

    arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"Termo_Perda_Ferias_{nome.replace(' ', '_')}.pdf", parent=tela.janela)
    if not arquivo: return
    
    try:
        doc = SimpleDocTemplate(arquivo, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=35, bottomMargin=35)
        elementos = []
        estilos = getSampleStyleSheet()
        
        est_data = ParagraphStyle('Data', parent=estilos['Normal'], fontSize=11, alignment=TA_RIGHT, spaceAfter=20)
        est_dest = ParagraphStyle('Destinatario', parent=estilos['Normal'], fontSize=11, alignment=TA_LEFT, spaceAfter=20)
        est_texto = ParagraphStyle('Texto', parent=estilos['Normal'], fontSize=11, alignment=TA_JUSTIFY, leading=16, spaceAfter=15)
        est_tabela_titulo = ParagraphStyle('TextoTabelaTit', parent=estilos['Normal'], fontSize=10, alignment=TA_CENTER, fontName="Helvetica-Bold")
        est_tabela_conteudo = ParagraphStyle('TextoTabelaCont', parent=estilos['Normal'], fontSize=10, alignment=TA_CENTER)
        
        elementos.append(Paragraph(f"Leme, {data_hoje_ext}.", est_data))
        elementos.append(Paragraph(f"Ilmo. Sr(a). <b>{nome}</b><br/>Nesta", est_dest))
        elementos.append(Paragraph("Prezado(a) Senhor(a) ", est_texto))
        
        txt_1 = f"Conforme determina o artigo 133 da CLT – Consolidação das Leis do Trabalho, face a V. S.ª ter percebido da Previdência Social prestações de Auxílio-Doença, Benefício(s) <b>{str_beneficios}</b>, o tempo em que esteve afastado, se maior que 06 (seis) meses dentro do período aquisitivo, perderá o direito a férias e iniciará nova contagem quando retornar ao trabalho."
        elementos.append(Paragraph(txt_1, est_texto))
        
        txt_2 = f"Assim, tendo se afastado, deixando de comparecer ao trabalho dentro do período aquisitivo de férias <b>{aq_de} a {aq_ate}</b>, temos abaixo a situação de suas férias:"
        elementos.append(Paragraph(txt_2, est_texto))
        
        elementos.append(Spacer(1, 5))
        
        dados_tabela = [
            [Paragraph("Período aquisitivo<br/>" + f"{aq_de} a {aq_ate}", est_tabela_titulo), 
             Paragraph("AFASTAMENTO<br/>(datas)", est_tabela_titulo), "", 
             Paragraph("Retorno<br/>Ao Serviço", est_tabela_titulo), 
             Paragraph("Início do novo período<br/>aquisitivo", est_tabela_titulo)],
            
            ["", Paragraph("Início", est_tabela_titulo), Paragraph("Fim", est_tabela_titulo), "", ""],
            
            [Paragraph(f"{aq_de}<br/>a<br/>{quinto_decimo_str}", est_tabela_conteudo),
             Paragraph(inicio_afast_str, est_tabela_conteudo), Paragraph(fim_afast_str, est_tabela_conteudo),
             Paragraph(ret_str, est_tabela_conteudo),
             Paragraph(f"{novo_aq_de} a {novo_aq_ate}", est_tabela_conteudo)],
             
            [create_brace(105), create_brace(135), "", create_brace(75), create_brace(145)],

            [Paragraph(str_dias_trabalhados, est_tabela_conteudo),
             Paragraph(str_tempo_afast_tbl, est_tabela_conteudo), "",
             Paragraph("Retorno ao<br/>trabalho", est_tabela_conteudo),
             Paragraph("12 meses a serem<br/>trabalhados", est_tabela_conteudo)],
             
            [create_brace(470), "", "", "", ""],

            [Paragraph(f"Perda do direito, e início de novo período em {novo_aq_de}.", ParagraphStyle('C', parent=est_tabela_conteudo, fontSize=11, fontName="Helvetica-Bold")), "", "", "", ""]
        ]
        
        t = Table(dados_tabela, colWidths=[110, 70, 70, 80, 150])
        t.setStyle(TableStyle([
            ('SPAN', (1,0), (2,0)), 
            ('SPAN', (0,0), (0,1)), 
            ('SPAN', (3,0), (3,1)), 
            ('SPAN', (4,0), (4,1)), 
            ('SPAN', (1,3), (2,3)), 
            ('SPAN', (1,4), (2,4)), 
            ('SPAN', (0,5), (4,5)), 
            ('SPAN', (0,6), (4,6)), 
            
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            
            ('GRID', (0,0), (-1,2), 0.5, colors.black),
            ('LINEBELOW', (0,2), (-1,2), 1, colors.black),
            
            ('LINEABOVE', (0,3), (-1,3), 0, colors.white),
            ('LINEBELOW', (0,3), (-1,6), 0, colors.white),
            ('LINEBEFORE', (0,3), (-1,6), 0, colors.white),
            ('LINEAFTER', (0,3), (-1,6), 0, colors.white),
            
            ('BOX', (0,0), (-1,-1), 1, colors.black),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,3), (-1,5), 0), 
            ('TOPPADDING', (0,3), (-1,5), 0),
        ]))
        elementos.append(t)
        
        elementos.append(Spacer(1, 15))
        
        txt_5 = f"Diante de seu retorno ao trabalho em {ret_str}, começa a fluir um novo período aquisitivo."
        elementos.append(Paragraph(txt_5, est_texto))
        
        elementos.append(Spacer(1, 35))
        elementos.append(Paragraph(f"<b>{empresa.upper()}</b>", ParagraphStyle('Empresa', alignment=TA_CENTER, fontName="Helvetica-Bold", fontSize=11)))
        elementos.append(Spacer(1, 35))
        elementos.append(Paragraph("Ciente", ParagraphStyle('Ciente', alignment=TA_CENTER, fontSize=11)))
        elementos.append(Spacer(1, 25))
        elementos.append(Paragraph("_____________________________________________________________", ParagraphStyle('Linha', alignment=TA_CENTER)))
        elementos.append(Paragraph(nome.upper(), ParagraphStyle('NomeAss', alignment=TA_CENTER, fontName="Helvetica-Bold", fontSize=11)))
        
        doc.build(elementos)
        abrir_pdf_sistema(arquivo)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar o Termo de Perda de Férias:\n{str(e)}", parent=tela.janela)


# ================= CASO 2: CÁRCERE =================

def gerar_termo_carcere(tela):
    ret_str = tela.ent_ret_data.get().strip()
    if not ret_str or len(ret_str) != 10:
        return messagebox.showwarning("Aviso", "Preencha a Data Efetiva de Retorno na aba 'Retorno' para emitir o termo de cárcere.", parent=tela.janela)

    empresa = tela.ent_empresa.get().strip()
    nome = tela.ent_nome.get().strip()
    aq_de = tela.ent_aq_de.get().strip()
    aq_ate = tela.ent_aq_ate.get().strip()
    periodo_de_str = tela.ent_per_de.get().strip()

    if not aq_de or not aq_ate or not periodo_de_str:
        return messagebox.showwarning("Aviso", "Certifique-se de que o Período Aquisitivo e o Início do Afastamento estão preenchidos.", parent=tela.janela)

    def calcular_meses_dias_comercial(dt_ini, dt_fim):
        if dt_fim < dt_ini:
            return 0, 0
        anos = dt_fim.year - dt_ini.year
        meses = dt_fim.month - dt_ini.month
        dias = dt_fim.day - dt_ini.day
        if dias < 0:
            meses -= 1
            dias += 30
        if meses < 0:
            anos -= 1
            meses += 12
        return (anos * 12 + meses), dias

    try:
        aq_de_dt = datetime.strptime(aq_de, "%d/%m/%Y").date()
        aq_ate_dt = datetime.strptime(aq_ate, "%d/%m/%Y").date()
        
        inicio_afast_dt = datetime.strptime(periodo_de_str, "%d/%m/%Y").date()
        ret_dt = datetime.strptime(ret_str, "%d/%m/%Y").date()
        fim_afast_dt = ret_dt - timedelta(days=1)
        
        ultimo_dia_trab_dt = inicio_afast_dt - timedelta(days=1)
        meses_antes, dias_antes = calcular_meses_dias_comercial(aq_de_dt, ultimo_dia_trab_dt)
        
        if dias_antes >= 30:
            meses_antes += dias_antes // 30
            dias_antes = dias_antes % 30

        if meses_antes > 0 and dias_antes == 0:
            str_antes = f"{meses_antes} meses<br/>trabalhados" if meses_antes > 1 else f"{meses_antes} mês<br/>trabalhado"
        elif meses_antes > 0 and dias_antes > 0:
            str_antes = f"{meses_antes} meses e {dias_antes:02d} dias<br/>trabalhados"
        else:
            str_antes = f"{dias_antes:02d} dias<br/>trabalhados"

        if dias_antes == 0:
            meses_falta = max(0, 12 - meses_antes)
            dias_falta = 0
        else:
            meses_falta = max(0, 11 - meses_antes)
            dias_falta = 30 - dias_antes

        if meses_falta > 0 and dias_falta == 0:
            str_falta = f"{meses_falta} meses<br/>trabalhados" if meses_falta > 1 else f"{meses_falta} mês<br/>trabalhado"
            str_falta_texto = f"{meses_falta:02d} meses de trabalho"
        elif meses_falta > 0 and dias_falta > 0:
            str_falta = f"{meses_falta} meses e {dias_falta:02d} dias<br/>trabalhados"
            str_falta_texto = f"{meses_falta:02d} meses e {dias_falta:02d} dias de trabalho"
        else:
            str_falta = f"{dias_falta:02d} dias<br/>trabalhados"
            str_falta_texto = f"{dias_falta:02d} dias de trabalho"

        base_ciclo_dt = ret_dt - timedelta(days=1)
        fim_meses_dt = somar_meses(base_ciclo_dt, meses_falta)
        fim_complementacao_dt = fim_meses_dt + timedelta(days=dias_falta)
        inicio_novo_aq_dt = fim_complementacao_dt + timedelta(days=1)

        tela.ent_novo_aq_de.delete(0, "end")
        tela.ent_novo_aq_de.insert(0, inicio_novo_aq_dt.strftime("%d/%m/%Y"))
        tela.atualizar_fim_novo_aquisitivo()
            
    except ValueError:
        return messagebox.showwarning("Erro", "Erro ao calcular datas. Verifique a formatação (DD/MM/AAAA).", parent=tela.janela)

    data_hoje_ext = data_por_extenso(datetime.now().date())
    arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"Termo_Carcere_{nome.replace(' ', '_')}.pdf", parent=tela.janela)
    if not arquivo: return
    
    try:
        doc = SimpleDocTemplate(arquivo, pagesize=A4, rightMargin=45, leftMargin=45, topMargin=35, bottomMargin=35)
        elementos = []
        estilos = getSampleStyleSheet()
        
        est_data = ParagraphStyle('Data', parent=estilos['Normal'], fontSize=11, alignment=TA_RIGHT, spaceAfter=20)
        est_dest = ParagraphStyle('Destinatario', parent=estilos['Normal'], fontSize=11, alignment=TA_LEFT, spaceAfter=20)
        est_texto = ParagraphStyle('Texto', parent=estilos['Normal'], fontSize=10, alignment=TA_JUSTIFY, leading=14, spaceAfter=15)
        est_tabela_titulo = ParagraphStyle('TextoTabelaTit', parent=estilos['Normal'], fontSize=9, alignment=TA_CENTER, fontName="Helvetica-Bold")
        est_tabela_conteudo = ParagraphStyle('TextoTabelaCont', parent=estilos['Normal'], fontSize=9, alignment=TA_CENTER)
        
        elementos.append(Paragraph(f"Leme, {data_hoje_ext}.", est_data))
        elementos.append(Paragraph(f"Ilmo. Sr. (a)<br/><b>{nome.upper()}</b><br/><u>Nesta</u>", est_dest))
        elementos.append(Paragraph("Prezado Senhor (a),", est_texto))
        
        txt_1 = "Conforme determina o artigo 131 da CLT_Consolidação das Leis do Trabalho, o tempo de trabalho anterior ao cárcere será computado no período aquisitivo das férias. Isto quer dizer que o período, durante o qual o empregado permanece afastado não é computado para efeito das férias, apenas o período anterior ao afastamento, e que será completado quando o empregado retornar ao serviço."
        elementos.append(Paragraph(txt_1, est_texto))
        
        txt_2 = f"Assim, tendo se afastado por cárcere, deixando de comparecer ao trabalho de {inicio_afast_dt.strftime('%d/%m/%Y')} a {fim_afast_dt.strftime('%d/%m/%Y')} temos abaixo a situação de suas férias:"
        elementos.append(Paragraph(txt_2, est_texto))
        
        elementos.append(Spacer(1, 10))
        
        dados_tabela = [
            # Linha 1: Cabeçalhos Principais
            [Paragraph(f"Período<br/>Aquisitivo<br/>{aq_de} a<br/>{aq_ate}", est_tabela_titulo), 
             Paragraph("Em cárcere (datas)", est_tabela_titulo), "", 
             Paragraph("Retorno<br/>Ao Serviço", est_tabela_titulo), 
             Paragraph("Complementação<br/>do período<br/>aquisitivo anterior<br/>ao afastamento", est_tabela_titulo),
             Paragraph("Início de Novo<br/>Período<br/>Aquisitivo", est_tabela_titulo)],
            
            # Linha 2: Sub-Cabeçalhos do Afastamento
            ["", Paragraph("de início", est_tabela_titulo), Paragraph("fim", est_tabela_titulo), "", "", ""],
            
            # Linha 3: As Datas
            [Paragraph(f"{aq_de}", est_tabela_conteudo),
             Paragraph(inicio_afast_dt.strftime("%d/%m/%Y"), est_tabela_conteudo), 
             Paragraph(fim_afast_dt.strftime("%d/%m/%Y"), est_tabela_conteudo),
             Paragraph(ret_str, est_tabela_conteudo),
             Paragraph(fim_complementacao_dt.strftime("%d/%m/%Y"), est_tabela_conteudo),
             Paragraph(inicio_novo_aq_dt.strftime("%d/%m/%Y"), est_tabela_conteudo)],
             
            # Linha 4: As Chaves Desenhadas
            [create_brace(80), create_brace(135), "", create_brace(65), create_brace(95), ""],

            # Linha 5: Os Textos Descritivos 
            [Paragraph(str_antes, est_tabela_conteudo),
             Paragraph("período não<br/>computado", est_tabela_conteudo), "",
             Paragraph("Retorno ao trabalho", est_tabela_conteudo),
             Paragraph(str_falta, est_tabela_conteudo),
             ""],
             
            # Linha 6: A Grande Chave Inferior
            [create_brace(390), "", "", "", "", ""],

            # Linha 7: Conclusão do Quadro
            [Paragraph("12 meses completos de trabalho", ParagraphStyle('C', parent=est_tabela_conteudo, fontName="Helvetica-Bold")), "", "", "", "", ""]
        ]
        
        t = Table(dados_tabela, colWidths=[85, 75, 75, 65, 105, 100])
        t.setStyle(TableStyle([
            ('SPAN', (1,0), (2,0)),
            ('SPAN', (0,0), (0,1)),
            ('SPAN', (3,0), (3,1)),
            ('SPAN', (4,0), (4,1)),
            ('SPAN', (5,0), (5,1)),
            ('SPAN', (1,3), (2,3)), 
            ('SPAN', (1,4), (2,4)), 
            ('SPAN', (0,5), (4,5)), 
            ('SPAN', (0,6), (4,6)), 
            ('SPAN', (5,2), (5,6)), 
            
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            
            ('GRID', (0,0), (-1,2), 0.5, colors.black),
            ('LINEBELOW', (0,2), (-1,2), 1, colors.black),
            
            ('LINEABOVE', (0,3), (4,3), 0, colors.white),
            ('LINEBELOW', (0,3), (4,6), 0, colors.white),
            ('LINEBEFORE', (0,3), (4,6), 0, colors.white),
            ('LINEAFTER', (0,3), (3,6), 0, colors.white),
            
            ('BOX', (0,0), (-1,-1), 1, colors.black),
            ('LINEBEFORE', (5,0), (5,-1), 1, colors.black),
            
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,3), (-1,5), 0), 
            ('TOPPADDING', (0,3), (-1,5), 0),
        ]))
        elementos.append(t)
        
        elementos.append(Spacer(1, 15))
        
        txt_conclusao = f"Diante de seu retorno ao trabalho em {ret_str} para completar o período aquisitivo das férias é necessário mais {str_falta_texto} até {fim_complementacao_dt.strftime('%d/%m/%Y')}, quando então estará com as férias completas, e a partir de {inicio_novo_aq_dt.strftime('%d/%m/%Y')} começa a fluir um novo período aquisitivo."
        elementos.append(Paragraph(txt_conclusao, est_texto))
        
        elementos.append(Spacer(1, 40))
        elementos.append(Paragraph(f"<b>{empresa.upper()}</b>", ParagraphStyle('Empresa', alignment=TA_CENTER, fontName="Helvetica-Bold", fontSize=10)))
        elementos.append(Spacer(1, 35))
        elementos.append(Paragraph("Ciente", ParagraphStyle('Ciente', alignment=TA_LEFT, fontSize=10)))
        elementos.append(Spacer(1, 20))
        elementos.append(Paragraph("_________________________________________________________", ParagraphStyle('Linha', alignment=TA_LEFT)))
        elementos.append(Paragraph(f"<b>{nome.upper()}</b>", ParagraphStyle('NomeAss', alignment=TA_LEFT, fontName="Helvetica-Bold", fontSize=10)))
        
        doc.build(elementos)
        abrir_pdf_sistema(arquivo)
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar o Termo de Cárcere:\n{str(e)}", parent=tela.janela)


# ================= CASO 3: SERVIÇO MILITAR =================

def gerar_termo_militar(tela):
    ret_str = tela.ent_ret_data.get().strip()
    if not ret_str or len(ret_str) != 10:
        return messagebox.showwarning("Aviso", "Preencha a Data Efetiva de Retorno na aba 'Retorno' para emitir o termo de serviço militar.", parent=tela.janela)

    empresa = tela.ent_empresa.get().strip()
    nome = tela.ent_nome.get().strip()
    aq_de = tela.ent_aq_de.get().strip()
    aq_ate = tela.ent_aq_ate.get().strip()
    periodo_de_str = tela.ent_per_de.get().strip()

    if not aq_de or not aq_ate or not periodo_de_str:
        return messagebox.showwarning("Aviso", "Certifique-se de que o Período Aquisitivo e o Início do Afastamento estão preenchidos.", parent=tela.janela)

    def calcular_meses_dias_comercial(dt_ini, dt_fim):
        if dt_fim < dt_ini:
            return 0, 0
        anos = dt_fim.year - dt_ini.year
        meses = dt_fim.month - dt_ini.month
        dias = dt_fim.day - dt_ini.day
        if dias < 0:
            meses -= 1
            dias += 30
        if meses < 0:
            anos -= 1
            meses += 12
        return (anos * 12 + meses), dias

    try:
        aq_de_dt = datetime.strptime(aq_de, "%d/%m/%Y").date()
        aq_ate_dt = datetime.strptime(aq_ate, "%d/%m/%Y").date()
        
        inicio_afast_dt = datetime.strptime(periodo_de_str, "%d/%m/%Y").date()
        ret_dt = datetime.strptime(ret_str, "%d/%m/%Y").date()
        fim_afast_dt = ret_dt - timedelta(days=1)
        
        ultimo_dia_trab_dt = inicio_afast_dt - timedelta(days=1)
        meses_antes, dias_antes = calcular_meses_dias_comercial(aq_de_dt, ultimo_dia_trab_dt)
        
        if dias_antes >= 30:
            meses_antes += dias_antes // 30
            dias_antes = dias_antes % 30

        if meses_antes > 0 and dias_antes == 0:
            str_antes = f"{meses_antes} meses<br/>trabalhados" if meses_antes > 1 else f"{meses_antes} mês<br/>trabalhado"
        elif meses_antes > 0 and dias_antes > 0:
            str_antes = f"{meses_antes} meses e {dias_antes:02d} dias<br/>trabalhados"
        else:
            str_antes = f"{dias_antes:02d} dias<br/>trabalhados"

        if dias_antes == 0:
            meses_falta = max(0, 12 - meses_antes)
            dias_falta = 0
        else:
            meses_falta = max(0, 11 - meses_antes)
            dias_falta = 30 - dias_antes

        if meses_falta > 0 and dias_falta == 0:
            str_falta = f"{meses_falta} meses<br/>trabalhados" if meses_falta > 1 else f"{meses_falta} mês<br/>trabalhado"
            str_falta_texto = f"{meses_falta:02d} meses de trabalho"
        elif meses_falta > 0 and dias_falta > 0:
            str_falta = f"{meses_falta} meses e {dias_falta:02d} dias<br/>trabalhados"
            str_falta_texto = f"{meses_falta:02d} meses e {dias_falta:02d} dias de trabalho"
        else:
            str_falta = f"{dias_falta:02d} dias<br/>trabalhados"
            str_falta_texto = f"{dias_falta:02d} dias de trabalho"

        base_ciclo_dt = ret_dt - timedelta(days=1)
        fim_meses_dt = somar_meses(base_ciclo_dt, meses_falta)
        fim_complementacao_dt = fim_meses_dt + timedelta(days=dias_falta)
        inicio_novo_aq_dt = fim_complementacao_dt + timedelta(days=1)

        tela.ent_novo_aq_de.delete(0, "end")
        tela.ent_novo_aq_de.insert(0, inicio_novo_aq_dt.strftime("%d/%m/%Y"))
        tela.atualizar_fim_novo_aquisitivo()
            
    except ValueError:
        return messagebox.showwarning("Erro", "Erro ao calcular datas. Verifique a formatação (DD/MM/AAAA).", parent=tela.janela)

    data_hoje_ext = data_por_extenso(datetime.now().date())
    arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"Termo_Servico_Militar_{nome.replace(' ', '_')}.pdf", parent=tela.janela)
    if not arquivo: return
    
    try:
        doc = SimpleDocTemplate(arquivo, pagesize=A4, rightMargin=45, leftMargin=45, topMargin=35, bottomMargin=35)
        elementos = []
        estilos = getSampleStyleSheet()
        
        est_data = ParagraphStyle('Data', parent=estilos['Normal'], fontSize=11, alignment=TA_RIGHT, spaceAfter=20)
        est_dest = ParagraphStyle('Destinatario', parent=estilos['Normal'], fontSize=11, alignment=TA_LEFT, spaceAfter=20)
        est_texto = ParagraphStyle('Texto', parent=estilos['Normal'], fontSize=10, alignment=TA_JUSTIFY, leading=14, spaceAfter=15)
        est_tabela_titulo = ParagraphStyle('TextoTabelaTit', parent=estilos['Normal'], fontSize=9, alignment=TA_CENTER, fontName="Helvetica-Bold")
        est_tabela_conteudo = ParagraphStyle('TextoTabelaCont', parent=estilos['Normal'], fontSize=9, alignment=TA_CENTER)
        
        elementos.append(Paragraph(f"Leme, {data_hoje_ext}.", est_data))
        elementos.append(Paragraph(f"Ilmo. Sr. (a)<br/><b>{nome.upper()}</b><br/><u>Nesta</u>", est_dest))
        elementos.append(Paragraph("Prezado Senhor (a),", est_texto))
        
        txt_1 = "Conforme determina o artigo 132 da CLT_Consolidação das Leis do Trabalho, o tempo de trabalho anterior ao serviço militar será computado no período aquisitivo das férias. Isto quer dizer que o período, durante o qual o empregado permanece afastado para o serviço militar não é computado para efeito das férias, apenas o período anterior ao afastamento, e que será completado quando o empregado retornar ao serviço."
        elementos.append(Paragraph(txt_1, est_texto))
        
        txt_2 = f"Assim, tendo se afastado para o serviço militar, deixando de comparecer ao trabalho de {inicio_afast_dt.strftime('%d/%m/%Y')} a {fim_afast_dt.strftime('%d/%m/%Y')} temos abaixo a situação de suas férias:"
        elementos.append(Paragraph(txt_2, est_texto))
        
        elementos.append(Spacer(1, 10))
        
        dados_tabela = [
            [Paragraph(f"Período<br/>Aquisitivo<br/>{aq_de} a<br/>{aq_ate}", est_tabela_titulo), 
             Paragraph("Afastamento p/ o Serviço Militar<br/>(datas)", est_tabela_titulo), "", 
             Paragraph("Retorno<br/>Ao Serviço", est_tabela_titulo), 
             Paragraph("Complementação<br/>do período<br/>aquisitivo anterior<br/>ao afastamento", est_tabela_titulo),
             Paragraph("Início de Novo<br/>Período<br/>Aquisitivo", est_tabela_titulo)],
            
            ["", Paragraph("de início", est_tabela_titulo), Paragraph("da baixa", est_tabela_titulo), "", "", ""],
            
            [Paragraph(f"{aq_de}", est_tabela_conteudo),
             Paragraph(inicio_afast_dt.strftime("%d/%m/%Y"), est_tabela_conteudo), 
             Paragraph(fim_afast_dt.strftime("%d/%m/%Y"), est_tabela_conteudo),
             Paragraph(ret_str, est_tabela_conteudo),
             Paragraph(fim_complementacao_dt.strftime("%d/%m/%Y"), est_tabela_conteudo),
             Paragraph(inicio_novo_aq_dt.strftime("%d/%m/%Y"), est_tabela_conteudo)],
             
            [create_brace(80), create_brace(135), "", create_brace(65), create_brace(95), ""],

            [Paragraph(str_antes, est_tabela_conteudo),
             Paragraph("período não<br/>computado", est_tabela_conteudo), "",
             Paragraph("dentro dos<br/>90 dias", est_tabela_conteudo),
             Paragraph(str_falta, est_tabela_conteudo),
             ""],
             
            [create_brace(390), "", "", "", "", ""],

            [Paragraph("12 meses completos de trabalho", ParagraphStyle('C', parent=est_tabela_conteudo, fontName="Helvetica-Bold")), "", "", "", "", ""]
        ]
        
        t = Table(dados_tabela, colWidths=[85, 75, 75, 65, 105, 100])
        t.setStyle(TableStyle([
            ('SPAN', (1,0), (2,0)),
            ('SPAN', (0,0), (0,1)),
            ('SPAN', (3,0), (3,1)),
            ('SPAN', (4,0), (4,1)),
            ('SPAN', (5,0), (5,1)),
            ('SPAN', (1,3), (2,3)), 
            ('SPAN', (1,4), (2,4)), 
            ('SPAN', (0,5), (4,5)), 
            ('SPAN', (0,6), (4,6)), 
            ('SPAN', (5,2), (5,6)), 
            
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            
            ('GRID', (0,0), (-1,2), 0.5, colors.black),
            ('LINEBELOW', (0,2), (-1,2), 1, colors.black),
            
            ('LINEABOVE', (0,3), (4,3), 0, colors.white),
            ('LINEBELOW', (0,3), (4,6), 0, colors.white),
            ('LINEBEFORE', (0,3), (4,6), 0, colors.white),
            ('LINEAFTER', (0,3), (3,6), 0, colors.white),
            
            ('BOX', (0,0), (-1,-1), 1, colors.black),
            ('LINEBEFORE', (5,0), (5,-1), 1, colors.black),
            
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,3), (-1,5), 0), 
            ('TOPPADDING', (0,3), (-1,5), 0),
        ]))
        elementos.append(t)
        
        elementos.append(Spacer(1, 15))
        
        txt_conclusao = f"Diante de seu retorno ao trabalho em {ret_str} para completar o período aquisitivo das férias é necessário mais {str_falta_texto} até {fim_complementacao_dt.strftime('%d/%m/%Y')}, quando então estará com as férias completas, e a partir de {inicio_novo_aq_dt.strftime('%d/%m/%Y')} começa a fluir um novo período aquisitivo."
        elementos.append(Paragraph(txt_conclusao, est_texto))
        
        elementos.append(Spacer(1, 40))
        elementos.append(Paragraph(f"<b>{empresa.upper()}</b>", ParagraphStyle('Empresa', alignment=TA_CENTER, fontName="Helvetica-Bold", fontSize=10)))
        elementos.append(Spacer(1, 35))
        elementos.append(Paragraph("Ciente", ParagraphStyle('Ciente', alignment=TA_LEFT, fontSize=10)))
        elementos.append(Spacer(1, 20))
        elementos.append(Paragraph("_________________________________________________________", ParagraphStyle('Linha', alignment=TA_LEFT)))
        elementos.append(Paragraph(f"<b>{nome.upper()}</b>", ParagraphStyle('NomeAss', alignment=TA_LEFT, fontName="Helvetica-Bold", fontSize=10)))
        
        doc.build(elementos)
        abrir_pdf_sistema(arquivo)
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar o Termo de Serviço Militar:\n{str(e)}", parent=tela.janela)


# ================= CASO 4: APOSENTADO (Afastamento por Motivo de Saúde) =================

def gerar_termo_aposentado(tela):
    ret_str = tela.ent_ret_data.get().strip()
    if not ret_str or len(ret_str) != 10:
        return messagebox.showwarning("Aviso", "Preencha a Data Efetiva de Retorno na aba 'Retorno' para emitir o termo de aposentado.", parent=tela.janela)

    empresa = tela.ent_empresa.get().strip()
    nome = tela.ent_nome.get().strip()
    aq_de = tela.ent_aq_de.get().strip()
    aq_ate = tela.ent_aq_ate.get().strip()
    periodo_de_str = tela.ent_per_de.get().strip()
    periodo_ate_str = tela.ent_per_ate.get().strip()

    if not aq_de or not aq_ate or not periodo_de_str:
        return messagebox.showwarning("Aviso", "Certifique-se de que o Período Aquisitivo e o Início do Afastamento estão preenchidos.", parent=tela.janela)

    def calcular_meses_dias_comercial(dt_ini, dt_fim):
        if dt_fim < dt_ini:
            return 0, 0
        anos = dt_fim.year - dt_ini.year
        meses = dt_fim.month - dt_ini.month
        dias = dt_fim.day - dt_ini.day
        if dias < 0:
            meses -= 1
            dias += 30
        if meses < 0:
            anos -= 1
            meses += 12
        return (anos * 12 + meses), dias

    try:
        aq_de_dt = datetime.strptime(aq_de, "%d/%m/%Y").date()
        aq_ate_dt = datetime.strptime(aq_ate, "%d/%m/%Y").date()
        
        inicio_afast_dt = datetime.strptime(periodo_de_str, "%d/%m/%Y").date()
        
        if periodo_ate_str and len(periodo_ate_str) == 10:
            fim_afast_dt = datetime.strptime(periodo_ate_str, "%d/%m/%Y").date()
        else:
            ret_dt_temp = datetime.strptime(ret_str, "%d/%m/%Y").date()
            fim_afast_dt = ret_dt_temp - timedelta(days=1)

        ret_dt = datetime.strptime(ret_str, "%d/%m/%Y").date()
        
        # 1. Dias Trabalhados ANTES do afastamento
        ultimo_dia_trab_dt = inicio_afast_dt - timedelta(days=1)
        meses_antes, dias_antes = calcular_meses_dias_comercial(aq_de_dt, ultimo_dia_trab_dt)
        
        if dias_antes >= 30:
            meses_antes += dias_antes // 30
            dias_antes = dias_antes % 30

        if meses_antes > 0 and dias_antes == 0:
            str_antes = f"{meses_antes} meses<br/>trabalhados" if meses_antes > 1 else f"{meses_antes} mês<br/>trabalhado"
        elif meses_antes > 0 and dias_antes > 0:
            str_antes = f"{meses_antes} meses e {dias_antes:02d} dias<br/>trabalhados"
        else:
            str_antes = f"{dias_antes:02d} dias<br/>trabalhados"

        # 2. Tempo que falta trabalhar para fechar 12 meses
        if dias_antes == 0:
            meses_falta = max(0, 12 - meses_antes)
            dias_falta = 0
        else:
            meses_falta = max(0, 11 - meses_antes)
            dias_falta = 30 - dias_antes

        if meses_falta > 0 and dias_falta == 0:
            str_falta = f"{meses_falta} meses<br/>trabalhados" if meses_falta > 1 else f"{meses_falta} mês<br/>trabalhado"
            str_falta_texto = f"{meses_falta:02d} meses de trabalho"
        elif meses_falta > 0 and dias_falta > 0:
            str_falta = f"{meses_falta} meses e {dias_falta:02d} dias<br/>trabalhados"
            str_falta_texto = f"{meses_falta:02d} meses e {dias_falta:02d} dias de trabalho"
        else:
            str_falta = f"{dias_falta:02d} dias<br/>trabalhados"
            str_falta_texto = f"{dias_falta:02d} dias de trabalho"

        # 3. Data final da complementação e início do novo período
        base_ciclo_dt = ret_dt - timedelta(days=1)
        fim_meses_dt = somar_meses(base_ciclo_dt, meses_falta)
        fim_complementacao_dt = fim_meses_dt + timedelta(days=dias_falta)
        inicio_novo_aq_dt = fim_complementacao_dt + timedelta(days=1)

        # Atualiza visualmente o campo na aba Retorno da tela (Sem popup!)
        tela.ent_novo_aq_de.config(state="normal")
        tela.ent_novo_aq_de.delete(0, "end")
        tela.ent_novo_aq_de.insert(0, inicio_novo_aq_dt.strftime("%d/%m/%Y"))
        tela.atualizar_fim_novo_aquisitivo()

    except ValueError:
        return messagebox.showwarning("Erro", "Erro ao processar as datas do afastamento.", parent=tela.janela)

    data_hoje_ext = data_por_extenso(datetime.now().date())
    arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"Termo_Aposentado_{nome.replace(' ', '_')}.pdf", parent=tela.janela)
    if not arquivo: return

    try:
        doc = SimpleDocTemplate(arquivo, pagesize=A4, rightMargin=45, leftMargin=45, topMargin=35, bottomMargin=35)
        elementos = []
        estilos = getSampleStyleSheet()
        
        est_data = ParagraphStyle('Data', parent=estilos['Normal'], fontSize=11, alignment=TA_RIGHT, spaceAfter=20)
        est_dest = ParagraphStyle('Destinatario', parent=estilos['Normal'], fontSize=11, alignment=TA_LEFT, spaceAfter=20)
        est_texto = ParagraphStyle('Texto', parent=estilos['Normal'], fontSize=10, alignment=TA_JUSTIFY, leading=14, spaceAfter=15)
        est_tabela_titulo = ParagraphStyle('TextoTabelaTit', parent=estilos['Normal'], fontSize=9, alignment=TA_CENTER, fontName="Helvetica-Bold")
        est_tabela_conteudo = ParagraphStyle('TextoTabelaCont', parent=estilos['Normal'], fontSize=9, alignment=TA_CENTER)
        
        elementos.append(Paragraph(f"Leme, {data_hoje_ext}.", est_data))
        elementos.append(Paragraph(f"Ilmo. Sr. (a)<br/><b>{nome.upper()}</b><br/><u>Nesta</u>", est_dest))
        elementos.append(Paragraph("Prezado Senhor (a),", est_texto))
        
        txt_1 = "Conforme determina o artigo 133 da CLT – Consolidação das Leis do Trabalho, o tempo de trabalho anterior ao afastamento por saúde será computado no período aquisitivo das férias. O empregado não teve oportunidade de perceber benefício do INSS por já ser aposentado, por este motivo ele não perde o direito a férias. Isto quer dizer que o período, durante o qual o empregado permanece afastado não é computado para efeito das férias, apenas o período anterior ao afastamento, e que será completado quando o empregado retornar ao serviço."
        elementos.append(Paragraph(txt_1, est_texto))
        
        txt_2 = f"Assim, tendo se afastado por motivo de saúde, deixando de comparecer ao trabalho de {inicio_afast_dt.strftime('%d/%m/%Y')} a {fim_afast_dt.strftime('%d/%m/%Y')} temos abaixo a situação de suas férias:"
        elementos.append(Paragraph(txt_2, est_texto))
        
        elementos.append(Spacer(1, 10))
        
        dados_tabela = [
            [Paragraph(f"Período Aquisitivo<br/><b>{aq_de}<br/>a<br/>{aq_ate}</b>", est_tabela_titulo), 
             Paragraph("Afastamento por motivo de<br/>saúde<br/>(datas)", est_tabela_titulo), "", 
             Paragraph("Retorno<br/>Ao Serviço", est_tabela_titulo), 
             Paragraph("Complementação<br/>do período<br/>aquisitivo anterior<br/>ao afastamento", est_tabela_titulo),
             Paragraph("Início de Novo<br/>Período<br/>Aquisitivo", est_tabela_titulo)],
            
            ["", Paragraph("de início", est_tabela_titulo), Paragraph("fim", est_tabela_titulo), "", "", ""],
            
            [Paragraph(f"{aq_de} a<br/>{ultimo_dia_trab_dt.strftime('%d/%m/%Y')}", est_tabela_conteudo),
             Paragraph(inicio_afast_dt.strftime('%d/%m/%Y'), est_tabela_conteudo), 
             Paragraph(fim_afast_dt.strftime('%d/%m/%Y'), est_tabela_conteudo),
             Paragraph(ret_str, est_tabela_conteudo),
             Paragraph(fim_complementacao_dt.strftime('%d/%m/%Y'), est_tabela_conteudo),
             Paragraph(inicio_novo_aq_dt.strftime('%d/%m/%Y'), est_tabela_conteudo)],
             
            [create_brace(80), create_brace(135), "", create_brace(65), create_brace(95), ""],

            [Paragraph(str_antes, est_tabela_conteudo),
             Paragraph("período não<br/>computado", est_tabela_conteudo), "",
             Paragraph("Retorno ao<br/>trabalho", est_tabela_conteudo),
             Paragraph(str_falta, est_tabela_conteudo),
             ""],
             
            [create_brace(390), "", "", "", "", ""],

            [Paragraph("12 meses completos de trabalho", ParagraphStyle('C', parent=est_tabela_conteudo, fontName="Helvetica-Bold")), "", "", "", "", ""]
        ]
        
        t = Table(dados_tabela, colWidths=[85, 75, 75, 65, 105, 100])
        t.setStyle(TableStyle([
            ('SPAN', (1,0), (2,0)),
            ('SPAN', (0,0), (0,1)),
            ('SPAN', (3,0), (3,1)),
            ('SPAN', (4,0), (4,1)),
            ('SPAN', (5,0), (5,1)),
            ('SPAN', (1,3), (2,3)), 
            ('SPAN', (1,4), (2,4)), 
            ('SPAN', (0,5), (4,5)), 
            ('SPAN', (0,6), (4,6)), 
            ('SPAN', (5,2), (5,6)), 
            
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            
            ('GRID', (0,0), (-1,2), 0.5, colors.black),
            ('LINEBELOW', (0,2), (-1,2), 1, colors.black),
            
            ('LINEABOVE', (0,3), (4,3), 0, colors.white),
            ('LINEBELOW', (0,3), (4,6), 0, colors.white),
            ('LINEBEFORE', (0,3), (4,6), 0, colors.white),
            ('LINEAFTER', (0,3), (3,6), 0, colors.white),
            
            ('BOX', (0,0), (-1,-1), 1, colors.black),
            ('LINEBEFORE', (5,0), (5,-1), 1, colors.black),
            
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,3), (-1,5), 0), 
            ('TOPPADDING', (0,3), (-1,5), 0),
        ]))
        elementos.append(t)
        
        elementos.append(Spacer(1, 15))
        
        txt_conclusao = f"Diante de seu retorno ao trabalho em {ret_str} para completar o período aquisitivo das férias é necessário mais {str_falta_texto} até {fim_complementacao_dt.strftime('%d/%m/%Y')}, quando então estará com as férias completas, e a partir de {inicio_novo_aq_dt.strftime('%d/%m/%Y')} começa a fluir um novo período aquisitivo."
        elementos.append(Paragraph(txt_conclusao, est_texto))
        
        elementos.append(Spacer(1, 40))
        elementos.append(Paragraph(f"<b>{empresa.upper()}</b>", ParagraphStyle('Empresa', alignment=TA_CENTER, fontName="Helvetica-Bold", fontSize=10)))
        elementos.append(Spacer(1, 35))
        elementos.append(Paragraph("Ciente", ParagraphStyle('Ciente', alignment=TA_LEFT, fontSize=10)))
        elementos.append(Spacer(1, 20))
        elementos.append(Paragraph("_________________________________________________________", ParagraphStyle('Linha', alignment=TA_LEFT)))
        elementos.append(Paragraph(f"<b>{nome.upper()}</b>", ParagraphStyle('NomeAss', alignment=TA_LEFT, fontName="Helvetica-Bold", fontSize=10)))
        
        doc.build(elementos)
        abrir_pdf_sistema(arquivo)
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar o Termo de Aposentado:\n{str(e)}", parent=tela.janela)