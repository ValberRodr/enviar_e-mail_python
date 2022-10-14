
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = ''
mail.Subject = 'Teste Python e-mail automatico'
mail.Body = 'Envio: Se você recebeu este e-mail, o python foi capaz de envia-lo automaticamente'
mail.HTMLBody = 'Se você recebeu este e-mail, o python foi capaz de envia-lo automaticamente' #this field is optional

# To attach a file to the email (optional):

attachment  = (r"")
mail.Attachments.Add(attachment)

mail.Send()
