# Análise 
Neste projeto, utilizei uma planilha em Excel contendo os dados de vendas de diversos shoppings. Através da análise dessas informações, desenvolvi um script no PyCharm utilizando a linguagem Python, com o auxílio das bibliotecas pandas e pywin32.

O objetivo principal foi importar os dados da planilha, identificar qual shopping obteve maior volume de vendas, calcular o faturamento total por loja, além do ticket médio por produto. Após a análise, o relatório foi formatado em HTML e enviado automaticamente por e-mail, utilizando uma conta do Gmail.
# Tecnologias Utilizadas
Python 3.x

PyCharm (IDE)

Pandas

Pywin32 (para integração com Outlook, opcional)

outlook e email (alternativa para envio via Gmail do win32)

Excel (.xlsx)
# Como Executar o Projeto
Clone o repositório:

bash
Copiar
Editar
git clone https://github.com/seu-usuario/seu-repositorio.git
Acesse o diretório do projeto:

bash
Copiar
Editar
cd nome-do-projeto
Crie e ative um ambiente virtual (opcional, mas recomendado):

bash
Copiar
Editar
python -m venv .venv
.venv\Scripts\activate
Instale as dependências:

bash
Copiar
Editar
pip install pandas pywin32
Coloque o arquivo Vendas.xlsx na mesma pasta do código.

Execute o script:

bash
Copiar
Editar
python nome_do_arquivo.py
# Observações Finais
Certifique-se de que o Outlook esteja instalado se usar pywin32.

O envio por Gmail é mais portátil e não depende de programas instalados.

A planilha deve conter as colunas ID Loja, Valor Final e Quantidade.

