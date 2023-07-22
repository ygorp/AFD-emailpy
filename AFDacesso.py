import datetime
import pyodbc
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Configurações de conexão ao banco de dados
server = 'nome servidor'
database = 'nome do banco'
username = 'usuario'
password = 'senha'
driver = '{SQL Server}'

# Configurações de e-mail
sender_email = 'email de origin'
receiver_email = 'email que recebe'
subject = 'Arquivo de dados de acessos'
body = 'Segue em anexo o arquivo com os dados de acessos.'

# Conectando ao banco de dados
conn = pyodbc.connect(
    f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'
)
cursor = conn.cursor()

# Obtendo a data atual
data_atual = datetime.date.today()

# Consulta SQL para obter os dados da tabela vw_acessos com filtro pela data corrente
sql_query = f"SELECT id, CONVERT(varchar(10), data, 120) AS data, hora, pessoa_n_folha FROM vw_acessos WHERE data = '{data_atual}' AND pessoa_n_folha IS NOT NULL"


try:
    # Executando a consulta SQL
    cursor.execute(sql_query)

    # Obtendo os resultados da consulta
    resultados = cursor.fetchall()

    # Gerando o arquivo de texto
    with open('dados_acessos.txt', 'w') as arquivo:
        for resultado in resultados:
            id = resultado[0]
            data = datetime.datetime.strptime(resultado[1], '%Y-%m-%d').strftime('%d%m%Y')  # Formata a data
            hora = datetime.datetime.strptime(resultado[2], '%H:%M:%S').strftime('%H%M%S')  # Formata a hora
            pessoa_n_folha = str(resultado[3]).zfill(11)
            
            # Escrevendo os dados no arquivo
            arquivo.write(f"{pessoa_n_folha}")
            arquivo.write(f"{data}")
            arquivo.write(f"{hora}\n")
            

    print("Arquivo de texto gerado com sucesso!")

    # Configurando a mensagem de e-mail
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    message.attach(MIMEText(body, 'plain'))

    # Anexando o arquivo ao e-mail
    attachment = open('dados_acessos.txt', 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= dados_acessos.txt")

    message.attach(part)

    #Enviando o e-mail
    server = smtplib.SMTP('smtp.outlook.com', 587)
    server.starttls()
    server.login(sender_email, 'babinogueira7292')  # Insira sua senha aqui

    server.send_message(message)
    server.quit()

    print("E-mail enviado com sucesso!")
    
except pyodbc.Error as e:
    print(f"Ocorreu um erro ao executar a consulta SQL: {e}")

# Fechando a conexão com o banco de dados
conn.close()
