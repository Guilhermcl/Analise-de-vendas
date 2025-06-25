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
# Análise da tabela
![image](https://github.com/user-attachments/assets/d92910c1-2df9-4fee-be24-03796ab34e61)
# Envio do gmail
![image](https://github.com/user-attachments/assets/872e61f7-eb2c-43f6-8cee-6950cd177c15)

# Observações Finais
Certifique-se de que o Outlook esteja instalado se usar pywin32.

O envio por Gmail é mais portátil e não depende de programas instalados.

A planilha deve conter as colunas ID Loja, Valor Final e Quantidade.

