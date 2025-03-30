import win32com.client as win32

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)

faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

# configurar as informações do seu e-mail
email.To = "destino; destino2"
email.Subject = "E-mail automático do Python"
email.HTMLBody = f"""
<p>Olá Lira, aqui é o código Python</p>

<p>O faturamento da loja foi de R${faturamento}</p>
<p>Vendemos {qtde_produtos} produtos</p>
<p>O ticket Médio foi de R${ticket_medio}</p>

<p>Abs,</p>
<p>Código Python</p>
"""

# anexo = "C://Users/joaop/Downloads/arquivo.xlsx"
# email.Attachments.Add(anexo)

email.Send()
print("Email Enviado")
import pyodbc

dados_conexao = (
    "Driver={SQL Server};"
    "Server=DESKTOP-T2JV7P5;"
    "Database=PythonSQL;"
)

conexao = pyodbc.connect(dados_conexao)
print("Conexão Bem Sucedida")

cursor = conexao.cursor()


id = 3
cliente = "Lira Python"
produto = "Carro"
data = "25/08/2021"
preco = 5000
quantidade = 1

comando = f"""INSERT INTO Vendas(id_venda, cliente, produto, data_venda, preco, quantidade)
VALUES
    ({id}, '{cliente}', '{produto}', '{data}', {preco}, {quantidade})"""

cursor.execute(comando)
cursor.commit()
janela = Tk()
janela.title("Cotação Atual de Moedas")
texto = Label(janela, text="Clique no botão para ver as cotações de moedas")
texto.grid(column=0, row=0, padx=10, pady=10)

botao = Button(janela, text="Buscar cotações", command=pegar_cotacoes)
botao.grid(column=0, row=1, padx=10, pady=10)

texto_resposta = Label(janela, text="")
texto_resposta.grid(column=0, row=2, padx=10, pady=10)


janela.mainloop()
git clone https://github.com/owner/repo.git
