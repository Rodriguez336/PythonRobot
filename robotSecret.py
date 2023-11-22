import win32com.client
import os

# Define o assunto do e-mail que você deseja salvar
assunto = "RES: Valide Seu Cliente - Erro de Preenchimento"

# Define o caminho da pasta de documentos no desktop
caminho = os.path.join(os.path.expanduser("A:\\TRANSFORMACAO DIGITAL\\1. app lei do bem\\9. chamados suporte app lei do bem\\testeRobot"))

# Cria a pasta de documentos se ela não existir
if not os.path.exists(caminho):
    os.makedirs(caminho)

# Conecta-se ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Obtém a pasta de entrada padrão
pasta_de_entrada = outlook.GetDefaultFolder(6)

# Obtém todos os e-mails na pasta de entrada
emails = pasta_de_entrada.Items

# Itera sobre cada e-mail
for e_mail in emails:
    # Verifica se o assunto do e-mail corresponde ao assunto especificado
    if e_mail.Subject == assunto:
        # Salva o e-mail na pasta de documentos
        nome_do_arquivo = os.path.join(caminho, f"{e_mail.Subject}.msg")
        e_mail.SaveAs(nome_do_arquivo)