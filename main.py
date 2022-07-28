import win32com.client as win32

if __name__ == "__main__":

    #fazendo integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    #contruindo e-mail
    email = outlook.CreateItem(0)

    #importar anexo  para o e-mail
    anexo1 = r"local_do_anexo"

    #definindo variáveis que serão inseridas no corpo do e-mail
    faturamento = 12500
    qtde_produtos = 840
    ticket_medio = faturamento / qtde_produtos
    #configurar as informações do seu e-mail
    email.To = "destino; destino2"
    email.CC = "destinoCC1; destinoCC2"
    email.Subject = "E-mail com Python"
    email.Attachments.Add(Source=anexo1)
    email.HTMLBody = f"""
    <p>Olá prezados!</p>

    <p>O faturamento fechado dessa semana foi de R$ {faturamento}</p>
    <p>Com um total de {qtde_produtos} produtos vendidos</p>
    <p>O ticket Médio foi de R$ {ticket_medio}</p>

    <p>Atenciosamente,</p>
    <p>Assinatura do E-mail</p>
    """

    email.Send()
    print("Email Enviado...")