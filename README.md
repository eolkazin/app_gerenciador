# Sistema de Gerenciamento de Obras

Este é um sistema desenvolvido em Python para gerenciar etapas de obras, permitindo o acompanhamento do progresso, cálculo de percentuais e geração de relatórios em PDF e Excel.

## Funcionalidades

- Cadastro de etapas de obras com informações como:
  - Nome da etapa
  - Responsável
  - Quantidade/Medida
  - Unidade de cálculo
  - Data de início
  - Duração em dias
  - Valor orçado
- Cálculo da porcentagem de conclusão da etapa em relação ao valor orçado.
- Atualização e salvamento das etapas diretamente em arquivos Excel, organizados por nome da obra.
- Geração automática de relatórios em PDF para cada obra, com detalhes de cada etapa.
- Integração para adicionar dados a planilhas existentes.
- Cálculo da data de conclusão de uma etapa com base na data de início e duração.

## Tecnologias Utilizadas

- **Python**: Linguagem de programação principal.
- **Tkinter**: Interface gráfica do usuário.
- **Pandas**: Manipulação de dados.
- **OpenPyXL**: Manipulação de arquivos Excel.
- **ReportLab**: Geração de PDFs.

## Requisitos

- Python 3.8 ou superior.
- Bibliotecas Python:
  - `tkinter`
  - `pandas`
  - `openpyxl`
  - `reportlab`
- Sistema Operacional: Windows.

## Instalação

1. Clone este repositório:
   ```bash
   git clone https://github.com/usuario/nome-do-repositorio.git
   ```

2. Navegue até o diretório do projeto:
   ```bash
   cd nome-do-repositorio
   ```

3. Instale os pacotes necessários:
   ```bash
   pip install -r requirements.txt
   ```

## Como Usar

1. Execute o programa principal:
   ```bash
   python main.py
   ```

2. Use a interface gráfica para:
   - Adicionar novas etapas.
   - Salvar etapas em arquivos Excel.
   - Gerar relatórios em PDF.
   - Calcular percentuais e datas de conclusão.

3. Os arquivos gerados serão organizados em pastas com o nome da obra.

## Estrutura do Projeto

```
|-- main.py                 # Arquivo principal do sistema.
|-- requirements.txt        # Dependências do projeto.
|-- /obras                  # Diretório contendo arquivos gerados (Excel e PDF).
|-- README.md               # Documentação do projeto.
```

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests.


**Desenvolvido com 💻 e ☕ por [Lucas Guerra].**
