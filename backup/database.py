import sqlite3
import os

DB_NAME = 'afastamentos.db'

def conectar_db():
    return sqlite3.connect(DB_NAME)

def inicializar_db():
    conn = conectar_db()
    cursor = conn.cursor()
    
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
            cod_empresa TEXT DEFAULT ''
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS atestados_historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_afastamento INTEGER,
            data_de TEXT,
            data_ate TEXT,
            qtd_dias INTEGER,
            observacoes TEXT,
            FOREIGN KEY(id_afastamento) REFERENCES afastamentos(id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cids_historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_afastamento INTEGER,
            codigo TEXT,
            descricao TEXT,
            FOREIGN KEY(id_afastamento) REFERENCES afastamentos(id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS analises_inss (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_afastamento INTEGER,
            data_analise TEXT,
            numero_analise TEXT,
            gerou_pericia TEXT,
            gerou_beneficio TEXT,
            numero_beneficio TEXT,
            situacao TEXT,
            concessao_de TEXT,
            concessao_ate TEXT,
            FOREIGN KEY(id_afastamento) REFERENCES afastamentos(id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pericias_inss (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_afastamento INTEGER,
            id_analise_mae INTEGER,
            data_agendamento TEXT,
            data_atendimento TEXT,
            especie TEXT,
            numero_beneficio TEXT,
            situacao TEXT,
            concessao_de TEXT,
            concessao_ate TEXT,
            FOREIGN KEY(id_afastamento) REFERENCES afastamentos(id)
        )
    ''')
    
    cursor.execute("PRAGMA table_info(afastamentos)")
    colunas = [col[1] for col in cursor.fetchall()]
    if 'observacoes' not in colunas: cursor.execute("ALTER TABLE afastamentos ADD COLUMN observacoes TEXT DEFAULT ''")
    if 'cpf' not in colunas: cursor.execute("ALTER TABLE afastamentos ADD COLUMN cpf TEXT DEFAULT ''")
    if 'domestica' not in colunas: cursor.execute("ALTER TABLE afastamentos ADD COLUMN domestica INTEGER DEFAULT 0")
    if 'cod_empresa' not in colunas: cursor.execute("ALTER TABLE afastamentos ADD COLUMN cod_empresa TEXT DEFAULT ''")
    
    # NOVAS COLUNAS PARA O CONTROLO DOS ALERTAS "FEITOS"
    if 'alerta_16_tratado' not in colunas: cursor.execute("ALTER TABLE afastamentos ADD COLUMN alerta_16_tratado INTEGER DEFAULT 0")
    if 'alerta_mat_tratado' not in colunas: cursor.execute("ALTER TABLE afastamentos ADD COLUMN alerta_mat_tratado INTEGER DEFAULT 0")

    cursor.execute("SELECT id, periodo_de, periodo_ate, qtd_dias FROM afastamentos")
    para_migrar = cursor.fetchall()
    for af in para_migrar:
        id_af, p_de, p_ate, q_dias = af
        if p_de and p_ate:
            cursor.execute("SELECT count(*) FROM atestados_historico WHERE id_afastamento=?", (id_af,))
            if cursor.fetchone()[0] == 0:
                cursor.execute("INSERT INTO atestados_historico (id_afastamento, data_de, data_ate, qtd_dias, observacoes) VALUES (?, ?, ?, ?, ?)", 
                               (id_af, p_de, p_ate, q_dias, 'Atestado Inicial (Migrado)'))

    cursor.execute("SELECT id, cid FROM afastamentos")
    para_migrar_cid = cursor.fetchall()
    for af in para_migrar_cid:
        id_af, cid_legado = af
        if cid_legado and str(cid_legado).strip() not in ("", "-"):
            cursor.execute("SELECT count(*) FROM cids_historico WHERE id_afastamento=?", (id_af,))
            if cursor.fetchone()[0] == 0:
                cursor.execute("INSERT INTO cids_historico (id_afastamento, codigo, descricao) VALUES (?, ?, ?)", 
                               (id_af, str(cid_legado).upper(), "CID Legado (Migração)"))

    conn.commit()
    conn.close()