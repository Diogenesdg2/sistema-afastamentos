# 🏥 Sistema de Controle de Afastamentos e Gestantes INSS

![Python](https://img.shields.io/badge/Python-3.x-blue.svg)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-lightgrey.svg)
![SQLite](https://img.shields.io/badge/Database-SQLite3-green.svg)
![ReportLab](https://img.shields.io/badge/PDF-ReportLab-red.svg)

Sistema desktop desenvolvido em Python para otimizar o fluxo do Departamento Pessoal e RH, focando no controle rigoroso de afastamentos previdenciários, acompanhamento de licenças-maternidade e gestão de recursos e contestações junto ao INSS.

O sistema elimina o controle manual em planilhas ao cruzar datas, gerar alertas de prazos críticos (como o 16º dia, risco de perda de férias e retornos), e automatizar a geração de termos de autorização e relatórios.

---

## ✨ Funcionalidades Principais

* **Gestão de Afastamentos:** Controle completo de atestados, motivos (Doença, Acidente, Maternidade, etc.), CIDs, e histórico de perícias médicas.
* **Acompanhamento de Gestantes:** Cálculo automatizado da projeção da 40ª semana de gestação e acompanhamento do período de estabilidade.
* **Gestão de Recursos e Contestações:** Módulo dedicado para gerir defesas, revisões e ações judiciais, mantendo histórico de andamentos e status (Aberto, Deferido, Indeferido).
* **Painel de Lembretes Inteligente (Resumo Matinal):** Pop-up automático na inicialização e sob demanda que cruza regras trabalhistas para alertar sobre:
  * 🚨 Prazo para agendamento de perícia (16º dia).
  * ⚠️ Risco de perda de direito a férias (Art. 133 da CLT - limite de 180 dias).
  * ⏳ Vencimento de prorrogações de benefícios.
  * 👶 Entrada na 37ª semana de gestação ou vencimento de licença-maternidade.
  * ⚖️ Processos de recurso sem movimentação há mais de 30 dias.
* **Geração de Documentos (PDF):** Emissão automatizada do *Termo de Autorização GOV.BR* com preenchimento exato e alinhamento visual de dados do funcionário.
* **Exportação e Relatórios:** Geração de relatórios filtrados em PDF e planilhas Excel (XLSX).
* **Filtros Dinâmicos:** Pesquisa rápida em tempo real por nome, empresa, tipos de afastamento e status ativo/baixado.

---

## 💻 Tecnologias e Bibliotecas Utilizadas

* **Python 3:** Linguagem base do projeto.
* **Tkinter & ttk:** Construção da Interface Gráfica de Usuário (GUI).
* **SQLite3:** Banco de dados relacional embutido e leve (sem necessidade de servidor externo).
* **ReportLab:** Geração dinâmica dos documentos e relatórios em PDF com formatação de precisão.
* **tkcalendar:** Componentes visuais para seleção e formatação de datas.
* **PyInstaller:** Compilação do código fonte para um executável `.exe` independente.

---

## 📂 Estrutura do Projeto

A arquitetura do projeto foi dividida em módulos lógicos para facilitar a manutenção e escalabilidade:

```text
afastamentos-inss/
├── app.py                     # Arquivo principal, dashboard, grids e lógica de alertas.
├── database.py                # Inicialização do banco, queries genéricas e migrações.
├── cadastro.py                # Tela de inserção de novos afastamentos.
├── gestante.py                # Tela de inserção de novas gestantes.
├── manutencao.py              # Tela principal de edição do afastamento e histórico de atestados/INSS.
├── manutencao_gestante.py     # Tela de edição e encerramento do ciclo da gestante.
├── manutencao_revisao.py      # Módulo isolado para gestão de Recursos, Revisões e Ações Judiciais.
├── relatorio.py               # Motor de exportação de dados para PDF e Excel.
├── importar.py                # Rotinas de importação de planilhas/dados externos.
├── requirements.txt           # Lista de dependências do projeto.
└── README.md                  # Documentação do projeto.
````

👤 Autor
Diógenes Daniel Gomes

Profissional de Tecnologia que trabalha em um escritório de contabilidade

GitHub: @Diogenesdg2

Desenvolvido como solução interna para modernização e automação de processos do Departamento Pessoal.

