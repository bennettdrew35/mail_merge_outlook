# -*- coding: utf-8 -*-
"""
Creloced on Wed Jun 13 11:18:36 2018

@author: drew.bennett
"""
import win32com.client as win32
import pandas as pd
import glob

# =============================================================================
# Combines datasets from EmailsPmPfParentsFqhcHf2018Q and OnpointPractProfilesDistList2017Sfy
# Creates new excel document to see information being pulled
# =============================================================================

excel_distro_list = r'\\nessie\public\Program_Initiatives\Blueprint_For_Health\Data_Team_Reporting\CHT\CHT Payments Spreadsheet\Email_Distribution.xlsx'
worksheet_distro_list = "Payer"
df = pd.read_excel(excel_distro_list, sheet_name=worksheet_distro_list)
# =============================================================================
# Runs for loop to create each of emails needed for the mail merge
# df.loc[df.index[i],#column name] represents the information in each cell
# =============================================================================
def mail_merge_info(e_To, e_CC=None, e_Subject="Hello", e_Body=None, e_attachment1=None, e_attachment2=None, e_attachment3=None):
    def attachment(loc):
        if loc is None:
            pass
        else:
            for x in glob.glob(loc):
                mail.attachments.Add(x)
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = e_To
    mail.CC = e_CC
    mail.Subject = e_Subject
    mail.Body = 'Message body'
    mail.HTMLBody =  e_Body
    attachment(e_attachment1)
    attachment(e_attachment2)
    attachment(e_attachment3)
    mail.Send()

for i in range(0, len(df.index)):
    print(i)
    path = r"\\nessie\public\Program_Initiatives\Blueprint_For_Health\Data_Team_Reporting\CHT\CHT Payments Spreadsheet\2018_Q3_Payment_Reports\SpecificByPayer\ChtPatientsAndPaymentsFor2018Q3_'"
    complete_path = path + df.loc[df.index[i],"Organization"] + "_*"
    html_data = "<html><body><p>Greetings,  "+ df.loc[df.index[i], 'Name'] +",</p></body></html>"
    loc_To = df.loc[df.index[i],'email_To']
    loc_CC = df.loc[df.index[i], 'email_CC']
    subject = df.loc[df.index[i],'Organization'] + ' CHT Patients and Payments for 2018-Q3'
    mail_merge_info(e_To=loc_To, e_CC=loc_CC, e_Subject=subject, e_Body=html_data, e_attachment1=complete_path)
