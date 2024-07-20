import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox

import langid
import win32com.client
from fpdf import FPDF, FontFace

# Global variables

heb_dict = {
    "title": "יצירת הודעות תזכורת ללקוחות",
    "email_prompt": ":כתובת אימייל",
    "pdf_name": ":pdf שם קובץ",
    "lower_bound": ":מספר ימים ללא יצירת קשר",
    "uper_bound": ":מספר ימים ללא הזמנה",
    "execute": "הרץ תכנית",
    "success": "!התכנית הושלמה",
    "failure": "!התכנית נכשלה",
    "success_msg": "!הודעות תזכורת נוצרו בהצלחה",
    "failure_msg": ".קרתה טעות במהלך הריצה\n.בדוק שכתובת המייל נכונה ושאאוטלוק פועל",
    "language_button": "English",
    "email_title": "!כתובת אימייל חסרה",
    "email": ".הכנס כתובת אימייל במקום המתאים ונסה שנית",
}
en_dict = {
    "title": "Creating Reminders for Customers",
    "email_prompt": "Email Address:",
    "pdf_name": "Name of PDF:",
    "lower_bound": "Minimum Days Without Contact:",
    "uper_bound": "Minimum Days Without Order:",
    "execute": "Run Program",
    "success": "Program Completed!",
    "failure": "Program Failed!",
    "success_msg": "Emails created Successfully!",
    "failure_msg": "An error ocurred running the program.\nCheck that the provided e-mail address is correct,\nand that Outlook is running.",
    "language_button": "עברית",
    "email_title": "Missing Address!",
    "email": "Provide an e-mail address in the entry slot and try again.",
}
text_dict = en_dict

customer_orders_dict = {}
sent_items_dict = {}
remove_list = set()
task_succeeded = True


# We define a main function first to make sure other definitions work correctly
def main(email, short, long):
    # get_global_vars()
    global task_succeeded

    # populate the dictionaries customer_orders_dict and sent_items_dict
    get_outlook_data(email)

    # Create tables
    if task_succeeded:
        table_1 = list(v[1] for v in customer_orders_dict.values())
        table_2 = list(v[1] for v in sent_items_dict.values())

        # Sort the tables by date from newest to oldest
        table_1.sort(key=lambda x: x[0], reverse=True)
        table_2.sort(key=lambda x: x[0], reverse=True)

        short_cutoff_date = datetime.now() - timedelta(days=short)
        short_cutoff_date = short_cutoff_date.strftime('%Y-%m-%d %H:%M:%S')
        long_cutoff_date = datetime.now() - timedelta(days=long)
        long_cutoff_date = long_cutoff_date.strftime('%Y-%m-%d %H:%M:%S')
        # Remove irrelevant items from the tables
        table_1 = remove_by_cutoff_date(table_1, table_2, short_cutoff_date)
        if remove_list != set():
            print_to_frame("removing customers...")
            table_1 = remove_customers_from_list(table_1, 1)
            table_2 = remove_customers_from_list(table_2, 2)
        
        # Convert the tables to pdf files
        print_to_pdf(table_1, "customer_orders")
        print_to_pdf(table_2, "sent_items")

        # Manage the email creations
        short_list, long_list = extract_and_sort_addresses(table_1, table_2, short_cutoff_date, long_cutoff_date)
        print_to_frame("Creating emails...")
        use_template_1to5(short_list)
        use_template_6to10(long_list)

# A function to help classify the language of a given text
def classify_language(text):
    for line in text.split('\n'):
        if langid.classify(line)[0] == 'he':
            return 'he'
    return langid.classify(text.split('\n')[0])

# Process data in batches for optimization
def process_data(mail_batch, folder_name, existing_domains):
    print_to_frame(f"Processing data from emails in {folder_name}...")
    # Initialize empty variables for domains
    to_domain, from_domain = '', ''

    for mail in mail_batch:
        # Extract relevant data (date, sender, to, cc, email_language, attachment, subject, to_domain/from_domain)
        # Skip emails that are replies or forwarded messages
        if mail.subject.startswith("Re:") or mail.subject.startswith("Fwd:"):
            continue
        try:
            date = ""
            if folder_name == "Sent Items":
                date = mail.SentOn.strftime('%Y-%m-%d %H:%M:%S')
            else:
                date = mail.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')
        except:
            print_to_frame("Something went wrong")
            continue
        sender = mail.SenderEmailAddress
        to = ''
        cc_list = []
        for recipient in mail.recipients:
            address = ''
            try:
                address = str(recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F"))
                if address:
                    if recipient.type == 1:
                        to = address
                    elif recipient.type == 2:
                        cc_list.append(address)
            except Exception as e:
                print_to_frame(f"\nError processing recipient: {e}")
                print_to_frame(f"Folder name: {folder_name}")
                print_to_frame(sender)
                continue
        cc = ' '.join(cc_list)
        email_lang = classify_language(mail.Body)
        email_language = 'English'
        if email_lang == 'he':
            email_language = 'תירבע'
        attachment = "Yes" if len(mail.Attachments) > 0 else "No"
        words = mail.subject.split(' ')
        for i in range(len(words)):
            word = words[i]
            if classify_language(word) == 'he':
                words[i] = word[::-1]
        subject = ' '.join(words)
        if classify_language(subject) == 'he':
            subject = ' '.join(words[::-1])
        
        # To differentiate between which folder we're working on, we check the name of the folder.
        # If "Sent Items", then we're currently working on sent items' folder, and need the to_domain data point
        # and we need to check if the domain exist in customer orders.
        if folder_name == "Sent Items":
            if '@' not in to:
                print_to_frame(f"\nMissing domain: {to}")
                continue
            to_domain = f"@{to.split('@')[1]}"
            # if domain doesn't appear in customer orders' folder, we don't want to add this item's data to the table
            if to_domain not in existing_domains:
                continue
        else:   # If folder_name is not "Sent Items", then we're currently working on customer orders' folder and we
                # need to get the from_domain data point as well as add the domain to our existing_domains set
            domain = mail.SenderEmailAddress
            if '@' not in domain:
                continue
            from_domain = f"@{domain.split('@')[1]}"
            existing_domains.add(from_domain)
        data = (date, sender, to, cc, email_language, attachment, subject, to_domain, from_domain, folder_name)
        # print(data)
        if folder_name == "Sent Items":
            if to not in sent_items_dict or sent_items_dict[to][0] < date:
                sent_items_dict[to] = [date, data]
        else:
            if sender not in customer_orders_dict or customer_orders_dict[sender][0] < date:
                customer_orders_dict[sender] = [date, data]
        from_domain, to_domain = '', ''

# Extract the data from Outlook folder into the global dictionaries
def extract_data(folder, folder_name, existing_domains=set()):
    print_to_frame(f"Extracting data from {folder_name}...")
    mails = list(folder.Items)

    # Split mails into batches
    batch_size = 100  # Adjust this value as needed
    mail_batches = [mails[i:i+batch_size] for i in range(0, len(mails), batch_size)]

    # Process each batch of mails
    for batch in mail_batches:
        process_data(batch, folder_name, existing_domains)

# Function to create the list of customers addresses to remove
def create_remove_list(remove_folder):
    for mail in remove_folder.Items:
        remove_list.add(mail.SenderEmailAddress)

# This function connects to outlook and finds the required folders
def get_outlook_data(email):
    global task_succeeded
    print_to_frame("Getting outlook data...")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  # connect to local outlook
    root_folder = outlook.Folders.Item(email) # Get the root of the folder hierarchy in your Outlook

    orders_folder_name = "Customer Orders"
    sent_folder_name = "Sent Items"
    remove_folder_name = "Remove"
    orders_folder, sent_folder = None, None

    # using try/except to open the required folders

    try:
        orders_folder = root_folder.Folders[orders_folder_name]
    except:
        print_to_frame(f"Folder '{orders_folder_name}' not found.")
        task_succeeded = False
        return
    try:
        sent_folder = root_folder.Folders[sent_folder_name]
    except:
        print_to_frame(f"Folder '{sent_folder_name}' not found.")
        task_succeeded = False
        return
    try:
        remove_folder = root_folder.Folders[remove_folder_name]
    except:
        print_to_frame(f"Folder '{remove_folder_name}' not found.")
        task_succeeded = False
        return

    # calling extract_data function to get the data into the tables
    existing_domains = set()
    extract_data(orders_folder, orders_folder_name, existing_domains)
    extract_data(sent_folder, sent_folder_name, existing_domains)
    create_remove_list(remove_folder)

# Remove items from the tables according to the cutoff date
def remove_by_cutoff_date(from_table, to_table, cutoff_date):
    addresses_to_remove = [tup[2] for tup in list(filter(lambda x: x[0] > cutoff_date, to_table))]
    return list(filter(lambda tup: tup[1] not in addresses_to_remove, from_table))

# Function to remove from the lists the customers that asked to be removed from mailing list
def remove_customers_from_list(table, col):
    new_table = []
    for line in table:
        if line[col] not in remove_list:
            new_table.append(line)
    return new_table

# The functions to create the email reminders
def use_template_1to5(address_list):
    print(address_list)

def use_template_6to10(address_list):
    print(address_list)

# Extract and sort email addresses from the created tables
def extract_and_sort_addresses(from_table, to_table, short_cutoff_date, long_cutoff_date):
    short_list = set()
    long_list = set()
    too_short = set()
    for line in from_table:
        received = line[0]
        address = line[1]
        if received < long_cutoff_date:
            long_list.add(address)
        elif received < short_cutoff_date:
            short_list.add(address)
        else:
            too_short.add(address)
    for line in to_table:
        received = line[0]
        address = line[2]
        if received < long_cutoff_date or address in too_short:
            continue
        elif received < short_cutoff_date:
            long_list.discard(address)
            short_list.add(address)
        else:
            long_list.discard(address)
            short_list.discard(address)
    return short_list, long_list

# Take the created data_table and create a pdf with a table containing the data for readable output
def print_to_pdf(data_table, pdf_name):
    print_to_frame(f"Attempting to create file {pdf_name}.pdf")
    if not data_table:  # If the table is empty, we return without creating the pdf
        return
    
    pdf = FPDF()    # Create pdf
    pdf.add_page('L')   # in landscape orientation
    pdf.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf')   # Add fonts that allow the pdf to print Hebrew characters
    pdf.add_font('DejaVu', 'B', 'DejaVuSansCondensed-Bold.ttf')
    pdf.set_font('DejaVu', '', size=7)     # Set the font to the font we just added
    # Open the table generator
    with pdf.table(width=260, col_widths=(25, 35, 35, 35, 20, 20, 30, 20, 20, 15), first_row_as_headings=False) as table:
        # Add header row
        table.row(cells=('Date', 'From', 'To', 'CC', 'Email Language', 'Attachment', 'Subject', 'To Domain', 'From Domain', 'Source'),style=FontFace(emphasis="Bold"))
        # Add the rest of the data to the table's rows
        for row in data_table:
            table.row(row)
    
    # Output the created pdf
    pdf.output(f"{pdf_name}.pdf")

#------------------------------------------------------------------------------------------------------------------------
#--------------------------------- Creating the GUI for the program -----------------------------------------------------
#========================================================================================================================

def arrange_widgets_in_grid():
    global text_dict
    language_button.grid(row=5, column=1, sticky='E')
    if text_dict == en_dict:
        email_label.grid(row=0, column=0, sticky='E')
        email_entry.grid(row=0, column=1, sticky='W')
        # pdf_label.grid(row=1, column=0, sticky='E')
        # pdf_entry.grid(row=1, column=1, sticky='W')
        lower_bound_label.grid(row=2, column=0, sticky='E')
        lower_bound_entry.grid(row=2, column=1, sticky='W')
        higher_bound_label.grid(row=3, column=0, sticky='E')
        higher_bound_entry.grid(row=3, column=1, sticky='W')
    else:
        email_label.grid(row=0, column=1, sticky='W')
        email_entry.grid(row=0, column=0, sticky='E')
        # pdf_label.grid(row=1, column=1, sticky='W')
        # pdf_entry.grid(row=1, column=0, sticky='E')
        lower_bound_label.grid(row=2, column=1, sticky='W')
        lower_bound_entry.grid(row=2, column=0, sticky='E')
        higher_bound_label.grid(row=3, column=1, sticky='W')
        higher_bound_entry.grid(row=3, column=0, sticky='E')
    process_button.grid(row=5, column=0, sticky='W')

def manage_widgets_labels():
    root.title(text_dict["title"])
    email_label.config(text=text_dict["email_prompt"])
    # pdf_label.config(text=text_dict["pdf_name"])
    lower_bound_label.config(text=text_dict["lower_bound"])
    higher_bound_label.config(text=text_dict["uper_bound"])
    process_button.config(text=text_dict["execute"])
    language_button.config(text=text_dict["language_button"])

def change_language():
    global text_dict
    if text_dict == en_dict:
        text_dict = heb_dict
    else:
        text_dict = en_dict
    manage_widgets_labels()
    arrange_widgets_in_grid()

def process_input():
    email = email_entry.get()
    if not email:
        messagebox.showinfo(text_dict["email_title"], text_dict["email"])
        return
    # pdf_filename = pdf_entry.get()
    lower_bound = int(lower_bound_entry.get())
    higher_bound = int(higher_bound_entry.get())

    start = datetime.now()
    main(email, lower_bound, higher_bound)
    end = datetime.now()
    elapsed = end - start

    # Display a message box to indicate completion
    if task_succeeded:
        messagebox.showinfo(text_dict["success"], text_dict["success_msg"])
    else:
        messagebox.showinfo(text_dict["failure"], text_dict["failure_msg"])
    print(f"Program finished in: {elapsed}")

def print_to_frame(text):
    print_text.insert(tk.INSERT,f"{text}\n")

# Create the main window
root = tk.Tk()
root.title(text_dict["title"])

# Create input fields with default values
email_label = tk.Label(root, text=text_dict["email_prompt"])
email_entry = tk.Entry(root)
lower_bound_label = tk.Label(root, text=text_dict["lower_bound"])
lower_bound_entry = tk.Entry(root)
lower_bound_entry.insert(0, "14")  # Set default lower bound
higher_bound_label = tk.Label(root, text=text_dict["uper_bound"])
higher_bound_entry = tk.Entry(root)
higher_bound_entry.insert(0, "30")  # Set default higher bound

# Create a button to process the input
process_button = tk.Button(root, text=text_dict["execute"], command=process_input)
language_button = tk.Button(root, text=text_dict["language_button"], command=change_language)

# Arrange widgets using grid layout
arrange_widgets_in_grid()
frame = tk.Frame(root, bg="#CCCCCC")
frame.grid(row=6, columnspan=2)
canvas = tk.Canvas(frame, bg="#CCCCCC")#, width=160, height=400)
canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
print_text = tk.Text(canvas, bg="white")
print_text.pack(side=tk.BOTTOM)

root.mainloop()
