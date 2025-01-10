import win32com.client as win32

def envia_email(nome, e):

    #integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    #email
    email = outlook.CreateItem(0)

    #configuração das informações do email
    email.To = e
    email.Subject = "teste1"
    email.HTMLBody = f"""
 <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f4f4f4; padding: 20px;">
        <tr>
            <td align="center">
                <table width="600px" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border: 1px solid #ddd; border-radius: 10px; overflow: hidden;">
                    <tr>
                        <td align="center" style="background-color: #090097; padding: 20px;">
                            <h1 style="color: white; font-size: 28px; margin: 0;">🎉 Feliz Aniversário, {nome}! 🎉</h1>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 20px; text-align: center;">
                            <p style="color: #090097; font-size: 18px; line-height: 1.6; margin: 0;">
                                Que este dia seja cheio de alegria, amor e celebrações! <br>
                        
                            </p>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #f4f4f4; text-align: center; padding: 10px;">
                            <p style="color: #090097; font-size: 14px; margin: 0;">
                                Enviado com 💖 para alegrar o seu dia!
                            </p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    """

    email.Send()
    print("EMAIL ENVIADO")



def main():
    nome = str()
    email = str()
    qtd = int()

    qtd = int(input("São quantos funcionários aniversariantes: "))
    for i in range(qtd):
        nome = input("Digite o nome: ")
        email = input("Digite o email: ")
        envia_email(nome, email)


if __name__ == "__main__":
    main()