import os
import shutil
from datetime import datetime
import xlwings as xw
import pandas as pd
import win32com.client as win32

# Define the source and destination directories
downloads_dir = os.path.expanduser('~\\Downloads')

parent_dir = os.path.dirname(os.getcwd())
# folder_name = datetime.now().strftime("%b%Y") # think of time between 2 months
folder_name = "May2023"
destination_dir = os.path.join(parent_dir, folder_name)

### think of what to do between different months

# Define excel directory and load excel file as df
excel_file_name = "ATO Correspondence Master 2023.xlsm"
excel_path = os.path.join(parent_dir, excel_file_name)

corr_df = pd.read_excel(excel_path, sheet_name = folder_name)
corr_df = corr_df[corr_df.Attended != "Y"] # only consider newly downloaded files
corr_df["Doc ID"] = corr_df["Doc ID"].astype(str) + ".pdf"

# Remove files with imp level = 0
files_to_remove = corr_df[corr_df["Importance Level"] == 0]["Doc ID"]

for file in files_to_remove:
    file_path = os.path.join(downloads_dir, file)
    os.remove(file_path)

# Move the rest to ATO corr folder by month
files_to_move = corr_df[corr_df['Importance Level'] != 0]['Doc ID']

for file in files_to_move:
    try: 
        src_path = os.path.join(downloads_dir, file)
        dst_path = os.path.join(destination_dir, file)
        shutil.move(src_path, dst_path)
    except FileNotFoundError as e:
        print("File not found:", e)

# Copy corr with imp level 1 and 2 to my folder
# my_folder = os.path.expanduser('~\\Desktop\\Phuong')
my_folder = "S:\\Staff\\Phuong"
files_to_copy = corr_df[(corr_df['Importance Level'] == 1) | 
                        (corr_df['Importance Level'] == 2)]['Doc ID']

for file in files_to_copy:
    try: 
        src_path = os.path.join(destination_dir, file)
        dst_path = os.path.join(my_folder, file)
        shutil.copy(src_path, dst_path)
    except FileNotFoundError as e:
        print("File not found:", e)

# Send corr with imp level of 3 and 4 to colleagues
to_sent_df = corr_df[(corr_df['Importance Level'] == 3) | (corr_df['Importance Level'] == 4)]
grouped_doc = to_sent_df.groupby("Email")["Doc ID"].agg(list)

# Set up the Outlook application
outlook = win32.Dispatch('outlook.application')
namespace = outlook.GetNamespace('MAPI')
mail_template_path = os.path.join(os.getcwd(), 'templates', 'ato_corr_template.msg')

for email, doc_list in grouped_doc.items():
    manager_name = email.split("@")[0]
    manager_name = manager_name[0].upper() + manager_name[1:]

    html_table = to_sent_df[to_sent_df["Email"] == email][["Name","Client ID","Subject","Doc ID"]]
    html_table = html_table.reset_index(drop=True).to_html()

    cc_list = ['daisy@mccanntax.com.au', 'phil@mccanntax.com.au', 'kevin@mccannfg.com.au','sonu@mccannfg.com.au']
    cc_list = [ele for ele in cc_list if ele != email]

    mail = outlook.CreateItemFromTemplate(mail_template_path)
    mail.Subject = 'ATO Correspondence from %s to %s' % (
        list(to_sent_df["Issue Date"])[1], list(to_sent_df["Issue Date"])[-1])
    
    # Replace placeholders in the email body with specific info
    mail.HTMLBody = mail.HTMLBody.replace('[COLLEAGUE]', manager_name)
    mail.HTMLBody = mail.HTMLBody.replace('[TABLE]', html_table)

    # Attach files
    for doc in doc_list:
        mail.Attachments.Add(os.path.join(destination_dir, doc))
    ### think of inserting a table with details of the attachments
    mail.To = email
    if email != "phuong@mccannfg.com.au":
        mail.CC = '; '.join(cc_list)

    mail.Send()

# # Ticked attended column and save
wb = xw.Book(excel_path)
sheet = wb.sheets[folder_name]
col_ind = sheet.range("1:1").value.index("Attended")+1
start_row = corr_df.index.min()+2
end_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

sheet.range((start_row,col_ind),(end_row,col_ind)).value = "Y"

wb.save()

# to bypass execution policy in 1 sesison
# PowerShell -ExecutionPolicy Bypass











