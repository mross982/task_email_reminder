
def emailerror(title, msg):
    import win32com.client
    from win32com.client import Dispatch, constants

    const=win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = title
    newMail.Body = msg
    newMail.To = "owners email address here"
    # newMail.CC = "Backup email address here"
    # attachment1 = r"E:\test\logo.png"

    # newMail.Attachments.Add(Source=attachment1)
    newMail.display()
    try:
        newMail.send()
    except:
        pass

def emailtask(msg, ls_e_addresses):
    import win32com.client
    from win32com.client import Dispatch, constants

    title = 'INFO: Common area 5S Checklist'

    const=win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = title
    newMail.Body = msg
    newMail.To = ls_e_addresses[0]
    newMail.CC = ls_e_addresses[1]
    # attachment1 = r"E:\test\logo.png"

    # newMail.Attachments.Add(Source=attachment1)
    newMail.display()
    try:
        newMail.send()
    except:
        pass