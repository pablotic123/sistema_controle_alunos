# Sistema de Controle de Alunos

![Python](https://img.shields.io/badge/Python-3.12-blue.svg) ![License](https://img.shields.io/badge/License-MIT-green.svg)

OBS: Texto do README.md gerado po IA

Bem-vindo ao **Sistema de Controle de Alunos**, uma aplicação desktop desenvolvida em Python com Tkinter para gerenciar instituições, cursos, turmas, professores e alunos. Este projeto utiliza uma interface gráfica intuitiva e um banco de dados SQLite para armazenar os dados, permitindo cadastro, consulta e exportação de relatórios em PDF, Excel e Word.

## Índice
- [Descrição](#descrição)
- [Funcionalidades](#funcionalidades)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Uso](#uso)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Contribuição](#contribuição)
- [Licença](#licença)
- [Contato](#contato)

## Descrição
O Sistema de Controle de Alunos é uma ferramenta projetada para facilitar a gestão acadêmica em instituições educacionais. Ele permite o cadastro e consulta de informações sobre instituições, cursos, turmas, professores e alunos, incluindo fotos e exportação de dados em diferentes formatos. O projeto é ideal para uso em escolas ou universidades que desejam uma solução simples e personalizável.

## Funcionalidades
- Cadastro e edição de instituições, cursos, turmas, professores e alunos.
- Consulta de registros com filtros dinâmicos.
- Visualização de um "Carômetro" (grade com fotos e nomes de alunos).
- Exportação de relatórios em PDF, Excel e Word.
- Interface gráfica amigável baseada em Tkinter.

## Pré-requisitos
Antes de instalar e executar o projeto, certifique-se de ter os seguintes itens instalados:
- **Python 3.12** ou superior (recomendado).
- Bibliotecas Python:
  - `tkinter` (geralmente incluído no Python padrão).
  - `pillow` (para manipulação de imagens).
  - `reportlab` (para exportação em PDF).
  - `openpyxl` (para exportação em Excel).
  - `python-docx` (para exportação em Word).
  - `sqlite3` (incluído no Python padrão).

## Instalação
1. Clone o repositório:
   git clone https://github.com/pablotic123/sistema_controle_alunos.git
   cd sistema_controle_alunos
2. Crie um ambiente virtual (opcional, mas recomendado):
   python -m venv venv
   venv\Scripts\activate
3. Instale as dependências:
   pip install pillow reportlab openpyxl python-docx
4. Execute o projeto:
   python main.py

## Uso
1. Ao iniciar o programa, você verá a tela inicial com estatísticas gerais.
2. Use o menu superior para:
   - Cadastros: Adicionar ou editar instituições, cursos, turmas, professores e alunos.
   - Consultas: Filtrar e visualizar registros.
   - Carômetro: Visualizar uma grade com fotos e nomes de alunos.
   - Exportar: Gerar relatórios em PDF, Excel ou Word.
3. Para adicionar fotos de alunos ou professores, use a opção "Selecionar" no cadastro e escolha uma imagem no formato JPG ou PNG.

## Estrutura do Projeto

sistema_controle_alunos/
│
├── main.py            # Ponto de entrada do programa
├── controller.py      # Lógica de controle da aplicação
├── view.py            # Interface gráfica com Tkinter
├── model.py           # Modelos de dados e interação com o banco
├── dados/             # Pasta para o banco de dados (ignorada no Git)
├── imagens/           # Pasta para fotos (ignorada no Git)
├── documentos/        # Pasta para relatórios gerados (ignorada no Git)
├── .gitignore         # Arquivo para ignorar arquivos desnecessários
└── README.md          # Este arquivo

## Contribuição
Contribuições são bem-vindas! Para contribuir:
1. Faça um fork do repositório.
2. Crie uma branch para sua funcionalidade:
   git checkout -b feature/nova-funcionalidade
3. Faça suas alterações e envie um commit:
   git commit -m "Adicionando nova funcionalidade"
4. Envie para o repositório remoto:
   git push origin feature/nova-funcionalidade
5. Abra um pull request no GitHub.

## Contato
Autor: Pablo Tic (pablotic123)
E-mail: pablotic123@gmail.com
GitHub: pablotic123
