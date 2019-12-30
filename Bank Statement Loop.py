import win32com.client

month="Dec19"
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
folder = outlook.folders['Month end WRL']
subfolder = folder[month]
subfolder=subfolder['Bank Statement']
subfoldermessages = subfolder.Items

email_list=[]
for email in subfoldermessages:
    if "7003" in email.Subject and "Report" in email.Subject:
        email_list.append(email)

attachments_list=[]
for email in email_list:
    for attachment in email.Attachments:
        attachments_list.append(attachment)

bankstatementfolder="//wisbis/dfs/Wesfarmers/GM reports/Resources Division/WDrive/Transition Folder/Accounting/Reporting/Month End/Divisional Office/Year Ended 30 June 2020/P06 WRL December 19/Bank Statements 7003"
for attachment in attachments_list:
    attachment.SaveAsFile(bankstatementfolder)








