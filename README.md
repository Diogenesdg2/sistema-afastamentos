🏢 Sistema de Gestão de Afastamentos e Férias
Um sistema desktop completo desenvolvido em Python para o departamento pessoal, focado no gerenciamento de afastamentos de colaboradores, controle de perícias do INSS e monitoramento automatizado do direito a férias (Art. 133 da CLT). Criado para otimizar rotinas de recursos humanos e contabilidade, o sistema garante segurança de dados e emissão de documentos legais com um clique.

🚀 Principais Funcionalidades
Gestão de Afastamentos: Cadastro e acompanhamento de múltiplos tipos de licenças (Doença, Acidente, Maternidade, Cárcere, Serviço Militar e Outros).

Termômetro de Férias (Art. 133 CLT): Monitoramento automático da soma de dias de benefício concedidos pelo INSS. O sistema alerta quando o colaborador atinge 180 dias de afastamento e projeta automaticamente o novo período aquisitivo.

Tratamento de Exceções Legais: Lógica dedicada para colaboradores Aposentados (que não perdem o direito a férias da mesma forma) e cálculos de estabilidade.

Central de Lembretes: Pop-ups matinais automáticos para alertar sobre o 16º dia de afastamento, fim de licença-maternidade e verificação de retorno ao trabalho.

Acompanhamento INSS: Módulo para gerir Análises Documentais e Perícias Médicas Presenciais, permitindo sincronizar a extensão dos atestados diretamente com os benefícios deferidos.

Geração Automática de Documentos: Emissão nativa de PDFs parametrizados, como o Termo de Perda de Férias e o Termo de Autorização de Acesso GOV.BR, prontos para assinatura.

🛠️ Tecnologias Utilizadas
Python 3.x: Linguagem principal da aplicação.

Tkinter & ttk: Construção da Interface Gráfica de Usuário (GUI).

tkcalendar: Componentes visuais para seleção intuitiva de datas.

SQLite3: Banco de dados relacional leve e embutido para persistência de registros.

ReportLab: Biblioteca para a geração e formatação avançada de documentos PDF.

📂 Estrutura do Projeto
app.py: Arquivo principal de inicialização e interface raiz do sistema.

manutencao.py: Módulo central de regras de negócio, interface de abas (Resumo, Acompanhamento, INSS, Retorno) e manipulação do banco de dados.

perda_ferias.py: Motor de cálculo de meses/dias comerciais e gerador dinâmico de PDFs usando ReportLab.

database.py: Gerenciamento de conexões, queries SQLite e cálculos centralizados de prazos.

🔒 Segurança e Privacidade de Dados (LGPD)
Este projeto lida com dados sensíveis de colaboradores (nomes, CPFs, laudos médicos, CIDs e senhas criptografadas). Por medidas rigorosas de segurança e adequação à LGPD, o arquivo do banco de dados (.db / .sqlite3) está estritamente ignorado no controle de versão via .gitignore.

Ao clonar este repositório, o sistema inicializará um banco de dados local novo e vazio na primeira execução. Nenhum dado real de empresa ou funcionário é exposto neste código-fonte.

