import customtkinter
import smtplib
import mimetypes
import logging
import time
import csv
import os
from tkinter import filedialog, messagebox
from tkhtmlview import HTMLLabel
from email.message import EmailMessage

# config do log
logging.basicConfig(
    filename='email_sending.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

frame = customtkinter.CTk()
frame.geometry("800x600")
frame.title("SMTP Server Bulk Sender")

contact_list = None
contact_count_label = None

def login(event = None):
    global server

    selected_value_str = combobox.get()
    smtp_port = int(selected_value_str)


    smtp_server = entry.get()    
    smtp_user = entry1.get()
    smtp_password = entry2.get()

    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(smtp_user, smtp_password)
        logging.info("Login successful")
        

        display_csv_frame()
    except Exception as e:
        label_result.configure(text="Login failed", text_color="red")
        logging.error(f"Login failed: {e}")

def clear_frame():
    for widget in frame.winfo_children():
        widget.destroy()

def display_csv_frame():
    clear_frame()
    global contact_count_label, contact_list, label_result2, alias_entry, from_entry

    label_main = customtkinter.CTkLabel(master=frame, text="Contact List", font=("Arial", 24, "bold"))
    label_main.pack(pady=20, padx=20)

    select_file_button = customtkinter.CTkButton(master=frame, text="Select CSV File", command=select_csv_file, font=("Arial", 14, "bold"))
    select_file_button.pack(pady=20, padx=20)

    alias_entry = customtkinter.CTkEntry(master=frame, placeholder_text="Alias", font=("Arial", 12))
    alias_entry.pack(pady=12, padx=10)
    
    from_entry = customtkinter.CTkEntry(master=frame, placeholder_text="(noreplay@domain.ext)", font=("Arial", 12))
    from_entry.pack(pady=12, padx=10)

    contact_count_label = customtkinter.CTkLabel(master=frame, text="Total Contacts: 0", font=("Arial", 14))
    contact_count_label.pack(pady=10, padx=10)

    contact_list = customtkinter.CTkTextbox(master=frame, height=10)
    contact_list.pack(pady=10, padx=10, fill='both', expand=True)
    contact_list.configure(state='normal')
    
    logout_button = customtkinter.CTkButton(master=frame, text="Logout", command=logout)
    logout_button.pack(pady=12, padx=10, side='left')

    next_button = customtkinter.CTkButton(master=frame, text="Next", command=display_message_frame)
    next_button.pack(pady=12, padx=10, side='right')

    label_result2 = customtkinter.CTkLabel(master=frame, text="", font=("Arial", 12))
    label_result2.pack(pady=12, padx=10)

def display_message_frame():
    if (contact_count != 0):
        global subject_entry, html_label, file_name_label, sender_alias, sender_email, testermail_entry, label_result_tester
        sender_alias = alias_entry.get()
        sender_email = from_entry.get()

        clear_frame()
        label_main = customtkinter.CTkLabel(master=frame, text="Message", font=("Arial", 24, "bold"))
        label_main.pack(pady=20, padx=20)

        subject_entry = customtkinter.CTkEntry(master=frame, placeholder_text="Subject", font=("Arial", 12))
        subject_entry.pack(pady=12, padx=10)

        select_file_button = customtkinter.CTkButton(master=frame, text="Select Attachment File", command=select_attachment, font=("Arial", 14, "bold"))
        select_file_button.pack(pady=20, padx=20)

        select_html_button = customtkinter.CTkButton(master=frame, text="Select HTML File as Body", command=select_html_file, font=("Arial", 14, "bold"))
        select_html_button.pack(pady=20, padx=20)

        testermail_entry = customtkinter.CTkEntry(master=frame, placeholder_text="Tester mail (opcional)", font=("Arial", 12))
        testermail_entry.pack()

        sendtester_button = customtkinter.CTkButton(master=frame, text="Send to tester", command=send_message_tester, font=("Arial", 14, "bold"))
        sendtester_button.pack()

        label_result_tester = customtkinter.CTkLabel(master=frame, text="", font=("Arial", 12))
        label_result_tester.pack()

        logout_button = customtkinter.CTkButton(master=frame, text="Logout", command=logout)
        logout_button.pack(pady=12, padx=10, side='left')

        next_button = customtkinter.CTkButton(master=frame, text="Send Mails", command=send_messages)
        next_button.pack(pady=12, padx=10, side='right')

        file_name_label = customtkinter.CTkLabel(master=frame, text=f"Loaded File: None")
        file_name_label.pack(pady=12, padx=10)
        
        initial_html_content = "<h1>Initial Content</h1><p>This is the initial HTML label.</p>"
        html_label = HTMLLabel(frame, html=initial_html_content)
        html_label.pack()
    else:
        label_result2.configure(text="You selected 0 contacts", text_color="red")
        logging.error(f"Trying to go next with 0 contacts")

def select_attachment():
    global att_file_path
    att_file_path = filedialog.askopenfilename(filetypes=[("All files", "*.*")])
    if att_file_path:
        display_file_name()

def display_file_name():
    file_name = os.path.basename(att_file_path)
    file_name_label.configure(text=f"Loaded File: {file_name}")
    

def select_html_file():
    html_file_path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
    if html_file_path:
        html_label.pack_forget()
        display_html(html_file_path)

def display_html(html_file_path):
    global html_content, html_label
    try:
        with open(html_file_path, 'r', encoding='utf-8') as file:
            html_content = file.read()
            

        # Create or update the HTMLLabel with the HTML content from the file
        html_label = HTMLLabel(frame, html=html_content)
        html_label.pack()
    
    except Exception as e:
        logging.error(f"Failed to load HTML file: {e}")
        messagebox.showerror("Error", f"Failed to load HTML file: {e}")


def select_csv_file():
    global csv_file_path
    csv_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if csv_file_path:
        display_contacts()

def find_email_column(first_row):
    for index, cell in enumerate(first_row):
        if '@' in cell:
            return index
    return None

def display_contacts():
    contact_list.configure(state='normal')
    contact_list.delete("1.0", customtkinter.END)
    contact_list.configure(state='disabled')
    try:
        with open(csv_file_path, mode='r') as csv_file:
            contact_list.configure(state='normal')
            csv_reader = csv.reader(csv_file)
            contacts = list(csv_reader)
            global email_column
            
            first_iteration = True
            email_column = None

            for row in contacts:
                if (first_iteration):
                    email_column = find_email_column(row)
                    first_iteration = False
                if (email_column != None):
                    email = row[email_column]
                    contact_list.insert('end', f"{email}\n")
                else:
                    messagebox.showerror("Error", f"Failed to read the CSV file: No email found")
                    logging.error(f"Failed to read the CSV file: No email found")
                    break
            global contact_count
            if (email_column != None):
                contact_list.configure(state='disabled')
                contact_count = len(contacts)
                contact_count_label.configure(text=f"Total Contacts: {contact_count}")
                label_result2.configure(text = "")
            else:
                contact_count = 0
                contact_list.configure(state='disabled')
                contact_count_label.configure(text=f"Total Contacts: {contact_count}")
                label_result2.configure(text = "")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to read the CSV file: {e}")
        logging.error(f"Failed to read the CSV file: {e}")

def logout():
    server.quit()
    clear_frame()
    display_login()

def display_login():
    global label_result, entry1, entry2, entry, combobox

    label = customtkinter.CTkLabel(master=frame, text="Login System", font=("Arial", 24, "bold"))
    label.pack(pady=12, padx=10)

    entry = customtkinter.CTkEntry(master=frame, placeholder_text="Server/Domain", font=("Arial", 12), width=200)
    entry.pack(pady=2, padx=10)
    entry.bind("<Return>", command=login)

    int_values = [25,26,465,587,2525,2526,3535]

    str_values = [str(i) for i in int_values]

    combobox = customtkinter.CTkComboBox(master=frame, font=("Arial", 12), width=200)
    combobox.set('465')
    combobox.configure(values = str_values)
    combobox.pack(pady=2)
    combobox.configure(state='readonly')

    entry1 = customtkinter.CTkEntry(master=frame, placeholder_text="Email/User", font=("Arial", 12), width=200)
    entry1.pack(pady=(11,1), padx=10)
    entry1.bind("<Return>", command=login)

    entry2 = customtkinter.CTkEntry(master=frame, placeholder_text="Password", show="*", font=("Arial", 12), width=200)
    entry2.pack(pady=2, padx=10)
    entry2.bind("<Return>", command=login)

    button = customtkinter.CTkButton(master=frame, text="Login", command=login, font=("Arial", 14, "bold"))
    button.pack(pady=12, padx=10)

    label_result = customtkinter.CTkLabel(master=frame, text="", font=("Arial", 12))
    label_result.pack(pady=20, padx=10)



def send_message_tester():
    subject = subject_entry.get()
    tester = testermail_entry.get()
    try:
        msg = EmailMessage()
        msg['To'] = tester

        msg['Subject'] = subject

        msg['From'] = f"{sender_alias} <{sender_email}>"

        msg.set_content(html_content, subtype='html')
        
        with open(att_file_path, 'rb') as f:
            file_data = f.read()

            mime_type, _ = mimetypes.guess_type(att_file_path)
            if mime_type is None:
                mime_type = 'application/octet-stream'

            maintype, subtype = mime_type.split('/', 1)
            msg.add_attachment(
                file_data,
                maintype=maintype,
                subtype=subtype,
                filename=os.path.basename(att_file_path)
            )
            server.send_message(msg)
            label_result_tester.configure(text="Sent successfully", text_color="green")
            logging.info(f"Email sent to {tester}")
    except Exception as e:
            label_result_tester.configure(text="Failed to send to tester", text_color="red")
            logging.error(f"Failed to send email to {tester}: {e}")    
        


# deixar mais bonito e ter feedback q enviou
def send_messages():
    subject = subject_entry.get()
    clear_frame()
    label = customtkinter.CTkLabel(master=frame, text="Log", font=("Arial", 24, "bold"))
    label.pack(pady=12, padx=10)

    log = customtkinter.CTkTextbox(master=frame, height=10)
    log.pack(pady=10, padx=10, fill='both', expand=True)
    log.configure(state='normal')

    with open(csv_file_path, mode='r') as csv_file:
        csv_reader = csv.reader(csv_file)

        # batch_size = 10
        # batch_count = 0
        
        for row in csv_reader:
            try:
                msg = EmailMessage()
                msg['To'] = row[email_column]

                msg['Subject'] = subject

                msg['From'] = f"{sender_alias} <{sender_email}>"


                msg.set_content(html_content, subtype='html')

                # attachment
                with open(att_file_path, 'rb') as f:
                    file_data = f.read()

                    mime_type, _ = mimetypes.guess_type(att_file_path)
                    if mime_type is None:
                        mime_type = 'application/octet-stream'

                    maintype, subtype = mime_type.split('/', 1)
                    msg.add_attachment(
                        file_data,
                        maintype=maintype,
                        subtype=subtype,
                        filename=os.path.basename(att_file_path)
                    )
                # ------

                server.send_message(msg)
                logging.info(f"Email sent to {row[1]}")

                # batch_count += 1
                # if batch_count >= batch_size:
                #     # se enviou ja 10 de uma vez
                #     logging.info("Batch limit reached, waiting for 60 seconds...")
                #     time.sleep(60)  # espera 60 segundos
                #     batch_count = 0
        
            except Exception as e:
                 logging.error(f"Failed to send email to {row[1]}: {e}")
        with open('email_sending.log', 'r', encoding='utf-8') as file:
            log_content = file.read()
            log.insert(customtkinter.END, log_content)


display_login()

frame.mainloop()
