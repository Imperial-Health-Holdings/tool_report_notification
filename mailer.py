import win32com.client as win32

def get_recipient_list(recipient:str) -> str:
    '''
    This function takes repicient email list from config_mailer file
    and process it into required format.

    Input from config_mailer file
        abc@imperialhealthholdings.com
        xyz@imperialhealthholdings.com
    Output from function
        abc@imperialhealthholdings.com; xyz@imperialhealthholdings.com
    '''
    recipient_list = [i.strip() for i in recipient.splitlines()]
    recipient_list = filter(None, recipient_list) # incase there's empty line, filter them out
    recipient_list = '; '.join(recipient_list)

    return recipient_list

def send_report_notification(recipient:str, subject:str, text:str) -> None:
    '''
    This function sends email using local Outlook account
    '''
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = get_recipient_list(recipient)
    mail.Subject = subject
    mail.Body = text
    #mail.HTMLBody = '<h2>HTML Message body</h2>'

    # To attach a file to the email (optional):
    if False:
        attachment  = "Path to the attachment"
        mail.Attachments.Add(attachment)

    mail.Send()

if __name__ == '__main__':
    None