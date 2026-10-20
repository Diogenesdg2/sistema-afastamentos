import sqlite3
import os
from datetime import datetime, timedelta

DB_NAME = 'afastamentos.db'

def conectar_db():
    return sqlite3.connect(DB_NAME)

def inicializar_db():
    conn = conectar_db()
    cursor = conn.cursor()
    
    # 1. TABELA PRINCIPAL DE AFASTAMENTOS
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS afastamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa TEXT, cod_empregado TEXT, nome_empregado TEXT,
            tipo_afastamento TEXT, ultimo_dia_trab TEXT, qtd_dias TEXT,
            periodo_de TEXT, periodo_ate TEXT, decimo_sexto_dia TEXT,
            cid TEXT, aposentado TEXT, data_obrigacao TEXT,
            obrigacao_concluida INTEGER DEFAULT 0, retorno_data TEXT,
            analise_distancia TEXT, atendimento_inss TEXT, pericia_presencial TEXT,
            concessao_beneficio TEXT, concessao_de TEXT, concessao_ate TEXT,
            recurso_revisao TEXT, nova_analise_distancia TEXT, novo_atendimento_inss TEXT,
            nova_pericia_data TEXT, nova_pericia_beneficio TEXT, nova_pericia_de TEXT,
            nova_pericia_ate TEXT, estabilidade TEXT, acidente_trab TEXT,
            email_enviado TEXT, retornou INTEGER DEFAULT 0,
            observacoes TEXT DEFAULT '', hiscre TEXT DEFAULT '',
            gov_br_autorizacao TEXT DEFAULT '', gov_br_senha TEXT DEFAULT '',
            situacao_pericia TEXT DEFAULT '', situacao_concessao TEXT DEFAULT '',
            cpf TEXT DEFAULT '', domestica INTEGER DEFAULT 0,
            cod_empresa TEXT DEFAULT '',
            alerta_16_tratado INTEGER DEFAULT 0,
            alerta_mat_tratado INTEGER DEFAULT 0,
            alerta_retorno_tratado INTEGER DEFAULT 0,
            aborto INTEGER DEFAULT 0,
            data_nascimento TEXT DEFAULT '',
            salario_familia TEXT DEFAULT '',
            salario_familia_obs TEXT DEFAULT '',
            imposto_renda TEXT DEFAULT '',
            imposto_renda_obs TEXT DEFAULT '',
            controlar_ferias INTEGER DEFAULT 0,
            per_aquisitivo_de TEXT DEFAULT '',
            per_aquisitivo_ate TEXT DEFAULT '',
            novo_per_aquisitivo_de TEXT DEFAULT '',
            novo_per_aquisitivo_ate TEXT DEFAULT '',
            alerta_ferias_tratado INTEGER DEFAULT 0
        )
    ''')

    # 2. TABELA: GESTANTES
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS gestantes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa TEXT, cod_empresa TEXT, cod_empregado TEXT, nome_empregado TEXT,
            semanas_exame REAL, data_exame TEXT, data_37_semanas TEXT,
            estabilidade TEXT, alerta_nasc_tratado INTEGER DEFAULT 0,
            status TEXT DEFAULT 'Gestante'
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS atestados_historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_afastamento INTEGER, data_de TEXT, data_ate TEXT,
            qtd_dias INTEGER, observacoes TEXT,
            FOREIGN KEY(id_afastamento) REFERENCES afastamentos(id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cids_historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_afastamento INTEGER, codigo TEXT, descricao TEXT,
            FOREIGN KEY(id_afastamento) REFERENCES afastamentos(id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS analises_inss (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_afastamento INTEGER, data_analise TEXT, numero_analise TEXT,
            gerou_pericia TEXT, gerou_beneficio TEXT, numero_beneficio TEXT,
            situacao TEXT, concessao_de TEXT, concessao_ate TEXT,
            prorrogacao_tratada INTEGER DEFAULT 0,
            FOREIGN KEY(id_afastamento) REFERENCES afastamentos(id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pericias_inss (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_afastamento INTEGER, id_analise_mae INTEGER,
            data_agendamento TEXT, data_atendimento TEXT, especie TEXT,
            numero_beneficio TEXT, situacao TEXT, concessao_de TEXT, concessao_ate TEXT,
            FOREIGN KEY(id_afastamento) REFERENCES afastamentos(id)
        )
    ''')
    
    # Migração automática segura de colunas
    cursor.execute("PRAGMA table_info(afastamentos)")
    colunas = [col[1] for col in cursor.fetchall()]
    
    novas_colunas = [
        ("observacoes", "TEXT DEFAULT ''"),
        ("cpf", "TEXT DEFAULT ''"),
        ("domestica", "INTEGER DEFAULT 0"),
        ("cod_empresa", "TEXT DEFAULT ''"),
        ("alerta_16_tratado", "INTEGER DEFAULT 0"),
        ("alerta_mat_tratado", "INTEGER DEFAULT 0"),
        ("alerta_retorno_tratado", "INTEGER DEFAULT 0"),
        ("aborto", "INTEGER DEFAULT 0"),
        ("data_nascimento", "TEXT DEFAULT ''"),
        ("salario_familia", "TEXT DEFAULT ''"),
        ("salario_familia_obs", "TEXT DEFAULT ''"),
        ("imposto_renda", "TEXT DEFAULT ''"),
        ("imposto_renda_obs", "TEXT DEFAULT ''"),
        ("controlar_ferias", "INTEGER DEFAULT 0"),
        ("per_aquisitivo_de", "TEXT DEFAULT ''"),
        ("per_aquisitivo_ate", "TEXT DEFAULT ''"),
        ("novo_per_aquisitivo_de", "TEXT DEFAULT ''"),
        ("novo_per_aquisitivo_ate", "TEXT DEFAULT ''"),
        ("alerta_ferias_tratado", "INTEGER DEFAULT 0")
    ]
    for nome_col, tipo_col in novas_colunas:
        if nome_col not in colunas:
            cursor.execute(f"ALTER TABLE afastamentos ADD COLUMN {nome_col} {tipo_col}")

    cursor.execute("PRAGMA table_info(analises_inss)")
    col_an = [col[1] for col in cursor.fetchall()]
    if 'prorrogacao_tratada' not in col_an:
        cursor.execute("ALTER TABLE analises_inss ADD COLUMN prorrogacao_tratada INTEGER DEFAULT 0")

    conn.commit()
    conn.close()


# =========================================================================
# ====== MOTOR INTELIGENTE DE PERÍODO AQUISITIVO E FÉRIAS (180 DIAS) ======
# =========================================================================

def calcular_fim_aquisitivo(dt_inicio_str):
    """Recebe DD/MM/AAAA e devolve o fim do período aquisitivo (+1 ano menos 1 dia)."""
    if not dt_inicio_str or len(dt_inicio_str.strip()) != 10:
        return ""
    try:
        dt_de = datetime.strptime(dt_inicio_str.strip(), "%d/%m/%Y")
        try:
            dt_um_ano = dt_de.replace(year=dt_de.year + 1)
        except ValueError:
            dt_um_ano = dt_de.replace(year=dt_de.year + 1, day=28)
        dt_fim = dt_um_ano - timedelta(days=1)
        return dt_fim.strftime("%d/%m/%Y")
    except ValueError:
        return ""

def obter_ids_mesmo_funcionario(cursor, cpf="", cod_empresa="", cod_empregado="", id_atual=None):
    """Localiza todos os IDs de afastamentos que pertencem ao mesmo colaborador."""
    ids = set()
    if id_atual:
        ids.add(int(id_atual))
        
    cpf_limpo = (cpf or "").strip()
    cod_emp_limpo = (cod_empresa or "").strip()
    cod_func_limpo = (cod_empregado or "").strip()

    if cpf_limpo:
        cursor.execute("SELECT id FROM afastamentos WHERE cpf=?", (cpf_limpo,))
        for r in cursor.fetchall(): ids.add(r[0])

    if cod_emp_limpo and cod_func_limpo:
        cursor.execute("SELECT id FROM afastamentos WHERE cod_empresa=? AND cod_empregado=?", (cod_emp_limpo, cod_func_limpo))
        for r in cursor.fetchall(): ids.add(r[0])

    return list(ids)

def buscar_dados_funcionario_existente(cpf="", cod_empresa="", cod_empregado=""):
    """Busca o registo mais recente do colaborador para autopreencher Nome e Período Aquisitivo."""
    conn = conectar_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cpf_limpo = (cpf or "").strip()
    cod_emp_limpo = (cod_empresa or "").strip()
    cod_func_limpo = (cod_empregado or "").strip()

    reg = None
    if cpf_limpo and len(cpf_limpo) >= 11:
        cursor.execute("SELECT * FROM afastamentos WHERE cpf=? ORDER BY id DESC LIMIT 1", (cpf_limpo,))
        reg = cursor.fetchone()
    
    if not reg and cod_emp_limpo and cod_func_limpo:
        cursor.execute("SELECT * FROM afastamentos WHERE cod_empresa=? AND cod_empregado=? ORDER BY id DESC LIMIT 1", (cod_emp_limpo, cod_func_limpo))
        reg = cursor.fetchone()

    conn.close()
    if not reg:
        return None

    # Se no último afastamento ele retornou e gerou um NOVO período aquisitivo (pós-perda de férias), puxa o novo!
    aq_de = reg['per_aquisitivo_de'] or ""
    aq_ate = reg['per_aquisitivo_ate'] or ""
    if reg['retornou'] == 1 and reg['novo_per_aquisitivo_de']:
        aq_de = reg['novo_per_aquisitivo_de']
        aq_ate = reg['novo_per_aquisitivo_ate'] or calcular_fim_aquisitivo(aq_de)

    return {
        'id': reg['id'],
        'empresa': reg['empresa'] or "",
        'cod_empresa': reg['cod_empresa'] or "",
        'nome_empregado': reg['nome_empregado'] or "",
        'cod_empregado': reg['cod_empregado'] or "",
        'cpf': reg['cpf'] or "",
        'controlar_ferias': reg['controlar_ferias'] if 'controlar_ferias' in reg.keys() else 0,
        'per_aquisitivo_de': aq_de,
        'per_aquisitivo_ate': aq_ate
    }

def calcular_dias_beneficio_aquisitivo(cpf="", cod_empresa="", cod_empregado="", aq_de_str="", aq_ate_str="", id_atual=None):
    """
    Soma APENAS os dias com situação 'Concedido' ou 'Concedida' (De / Até) nas tabelas
    analises_inss e pericias_inss de TODOS os afastamentos do colaborador que caiam dentro
    do Período Aquisitivo [aq_de_str, aq_ate_str], sem duplicar dias sobrepostos.
    """
    if not aq_de_str or not aq_ate_str:
        return 0, 0
    try:
        aq_de = datetime.strptime(aq_de_str.strip(), "%d/%m/%Y").date()
        aq_ate = datetime.strptime(aq_ate_str.strip(), "%d/%m/%Y").date()
    except ValueError:
        return 0, 0

    conn = conectar_db()
    cursor = conn.cursor()

    ids_afast = obter_ids_mesmo_funcionario(cursor, cpf, cod_empresa, cod_empregado, id_atual)
    if not ids_afast:
        conn.close()
        return 0, 0

    qmarks = ",".join("?" for _ in ids_afast)
    periodos_concedidos = []

    # 1. Busca nas Análises Documentais Concedidas
    cursor.execute(f"""
        SELECT concessao_de, concessao_ate FROM analises_inss 
        WHERE id_afastamento IN ({qmarks}) AND situacao IN ('Concedido', 'Concedida')
    """, ids_afast)
    for r in cursor.fetchall():
        if r[0] and r[1]: periodos_concedidos.append((r[0], r[1]))

    # 2. Busca nas Perícias Presenciais Concedidas
    cursor.execute(f"""
        SELECT concessao_de, concessao_ate FROM pericias_inss 
        WHERE id_afastamento IN ({qmarks}) AND situacao IN ('Concedido', 'Concedida')
    """, ids_afast)
    for r in cursor.fetchall():
        if r[0] and r[1]: periodos_concedidos.append((r[0], r[1]))

    conn.close()

    dias_unicos = set()
    for ini_s, fim_s in periodos_concedidos:
        try:
            dt_ini = datetime.strptime(ini_s.strip(), "%d/%m/%Y").date()
            dt_fim = datetime.strptime(fim_s.strip(), "%d/%m/%Y").date()
            # Recorta estritamente os dias que caem dentro do Período Aquisitivo
            ini_efetivo = max(dt_ini, aq_de)
            fim_efetivo = min(dt_fim, aq_ate)
            if ini_efetivo <= fim_efetivo:
                atual = ini_efetivo
                while atual <= fim_efetivo:
                    dias_unicos.add(atual)
                    atual += timedelta(days=1)
        except ValueError:
            pass

    return len(dias_unicos), len(ids_afast)

def sincronizar_aquisitivo_funcionario(cpf, cod_empresa, cod_empregado, controlar, aq_de, aq_ate):
    """Sincroniza o Período Aquisitivo atual entre os afastamentos em aberto do mesmo funcionário."""
    conn = conectar_db()
    cursor = conn.cursor()
    ids = obter_ids_mesmo_funcionario(cursor, cpf, cod_empresa, cod_empregado)
    if ids:
        qmarks = ",".join("?" for _ in ids)
        cursor.execute(f"""
            UPDATE afastamentos 
            SET controlar_ferias=?, per_aquisitivo_de=?, per_aquisitivo_ate=?
            WHERE id IN ({qmarks}) AND retornou=0
        """, [controlar, aq_de, aq_ate] + ids)
        conn.commit()
    conn.close()