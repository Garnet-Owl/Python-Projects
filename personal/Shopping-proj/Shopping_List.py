import tkinter as tk
from tkinter import filedialog, messagebox
import datetime
import os
import pandas as pd
import csv
import docx
from PyPDF2 import PdfReader
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Global variables
loaded_file_path = None
receiver_email = None
file_path = None  # Holds the path of the saved file

def create_shopping_lists_folder():
    if not os.path.exists('../Shopping Lists'):
        os.makedirs('../Shopping Lists')

def add_row():
    row = len(entries) + 1
    item_entry = tk.Entry(scrollable_frame_inner)
    item_entry.grid(row=row, column=0)
    price_entry = tk.Entry(scrollable_frame_inner)
    price_entry.grid(row=row, column=1)
    quantity_entry = tk.Entry(scrollable_frame_inner)
    quantity_entry.grid(row=row, column=2)
    entries.append((item_entry, price_entry, quantity_entry))

def calculate_total():
    global file_path
    total = 0
    for item_entry, price_entry, quantity_entry in entries:
        try:
            price = float(price_entry.get())
            quantity = int(quantity_entry.get())
            total += price * quantity
        except ValueError:
            continue
    total_label.config(text=f"Total: Kshs. {total}")
    save_list(total)

def save_list(total):
    global loaded_file_path, file_path

    # Ensure the folder exists
    create_shopping_lists_folder()

    file_format = file_format_var.get()

    if loaded_file_path:
        # Save to the loaded file
        save_path = loaded_file_path
        last_edited_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    else:
        # Create a new file with the current date and time
        current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        if file_format == "Text":
            save_path = os.path.join("../Shopping Lists", f"shopping_list_{current_time}.txt")
        elif file_format == "CSV":
            save_path = os.path.join("../Shopping Lists", f"shopping_list_{current_time}.csv")
        last_edited_time = None

    file_path = save_path  # Set the global file_path

    try:
        if file_format == "Text":
            with open(save_path, 'w') as f:
                f.write("Shopping List\n")
                f.write("-----------------------------\n")
                f.write("Item: Price: Quantity\n")
                f.write("-----------------------------\n")
                for item_entry, price_entry, quantity_entry in entries:
                    item = item_entry.get().strip()
                    price = price_entry.get().strip()
                    quantity = quantity_entry.get().strip()
                    if item and price and quantity:
                        f.write(f"{item}: {price}: {quantity}\n")
                f.write("-----------------------------\n")
                f.write(f"Total: Kshs. {total}\n")
                if last_edited_time:
                    f.write(f"\nLast edited: {last_edited_time}\n")
        elif file_format == "CSV":
            with open(save_path, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Item", "Price (Kshs)", "Quantity"])
                for item_entry, price_entry, quantity_entry in entries:
                    item = item_entry.get().strip()
                    price = price_entry.get().strip()
                    quantity = quantity_entry.get().strip()
                    if item and price and quantity:
                        writer.writerow([item, price, quantity])
                writer.writerow(["Total", f"Kshs. {total}", ""])
                if last_edited_time:
                    writer.writerow(["Last edited", last_edited_time, ""])
        messagebox.showinfo("Save Successful", f"List saved as {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save file: {str(e)}")

def load_list():
    global loaded_file_path
    file_path = filedialog.askopenfilename(
        filetypes=[("Text files", "*.txt"), ("Word files", "*.docx"), ("PDF files", "*.pdf"), ("Excel files", "*.xlsx")]
    )
    if not file_path:
        return

    loaded_file_path = file_path

    try:
        if file_path.endswith('.txt'):
            with open(file_path, 'r') as f:
                lines = f.readlines()
        elif file_path.endswith('.docx'):
            doc = docx.Document(file_path)
            lines = [para.text for para in doc.paragraphs]
        elif file_path.endswith('.pdf'):
            reader = PdfReader(file_path)
            lines = []
            for page in reader.pages:
                lines.extend(page.extract_text().split('\n'))
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, header=None)
            lines = [f"{row[0]}: {row[1]}: {row[2]}" for row in df.values]

        parse_loaded_data(lines)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {str(e)}")

def parse_loaded_data(lines):
    # Clear existing entries
    for item_entry, price_entry, quantity_entry in entries:
        item_entry.destroy()
        price_entry.destroy()
        quantity_entry.destroy()
    entries.clear()

    total_found = False
    last_edited_time = None

    for line in lines:
        line = line.strip()
        if line.startswith("Item:"):
            continue
        if ":" in line:
            if "Total:" in line:
                total_found = True
                total = float(line.split(":")[1].strip().replace("Kshs.", "").strip())
                total_label.config(text=f"Total: Kshs. {total}")
                continue
            if "Last edited:" in line:
                last_edited_time = line.split(":")[1].strip()
                continue
            parts = line.split(":")
            if len(parts) == 3:
                item, price, quantity = parts
                add_row()
                entries[-1][0].insert(0, item.strip())
                entries[-1][1].insert(0, price.strip())
                entries[-1][2].insert(0, quantity.strip())

def send_email():
    global receiver_email, file_path
    if not file_path or not receiver_email:
        messagebox.showerror("Error", "No file path or receiver email specified.")
        return

    sender_email = "james544von@gmail.com"
    sender_password = "#Dallas_HT_2024_VK"

    subject = "Shopping List"
    body = f"Attached is the shopping list.\n\nTotal: Kshs. {total_label.cget('text').split(' ')[1]}\n\nBest regards,\nYour Shopping List App"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(file_path, 'r') as attachment:
        part = MIMEText(attachment.read(), 'plain')
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
        msg.attach(part)

    try:
        with smtplib.SMTP('smtp.example.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        messagebox.showinfo("Email Sent", f"Shopping list emailed to {receiver_email}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send email: {str(e)}")

# Initialize the tkinter window
root = tk.Tk()
root.title("Shopping List Calculator")

# Create a scrollable frame
scrollable_frame = tk.Frame(root)
scrollable_frame.pack(fill=tk.BOTH, expand=True)

canvas = tk.Canvas(scrollable_frame)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(scrollable_frame, orient="vertical", command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas.configure(yscrollcommand=scrollbar.set)

scrollable_frame_inner = tk.Frame(canvas)
canvas.create_window((0, 0), window=scrollable_frame_inner, anchor="nw")

def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

scrollable_frame_inner.bind("<Configure>", on_frame_configure)

# Header labels
tk.Label(scrollable_frame_inner, text="Item").grid(row=0, column=0)
tk.Label(scrollable_frame_inner, text="Price (Kshs)").grid(row=0, column=1)
tk.Label(scrollable_frame_inner, text="Quantity").grid(row=0, column=2)

# List to store the entry widgets
entries = []

# Add initial row
add_row()

# Buttons and options
button_frame = tk.Frame(root)
button_frame.pack(fill=tk.X)

add_row_button = tk.Button(button_frame, text="Add Row", command=add_row)
add_row_button.pack(side=tk.LEFT)

calculate_button = tk.Button(button_frame, text="Calculate Total", command=calculate_total)
calculate_button.pack(side=tk.LEFT)

load_button = tk.Button(button_frame, text="Load List", command=load_list)
load_button.pack(side=tk.LEFT)

email_button = tk.Button(button_frame, text="Send Email", command=send_email)
email_button.pack(side=tk.LEFT)

# Save file format options
file_format_var = tk.StringVar(value="Text")
file_format_frame = tk.Frame(root)
file_format_frame.pack()

tk.Radiobutton(file_format_frame, text="Text", variable=file_format_var, value="Text").pack(side=tk.LEFT)
tk.Radiobutton(file_format_frame, text="CSV", variable=file_format_var, value="CSV").pack(side=tk.LEFT)

# Entry for receiver email
email_label = tk.Label(root, text="Receiver Email:")
email_label.pack()

email_entry = tk.Entry(root)
email_entry.pack()

def update_receiver_email(event=None):
    global receiver_email
    receiver_email = email_entry.get()

email_entry.bind("<FocusOut>", update_receiver_email)

# Label to display the total
total_label = tk.Label(root, text="Total: Kshs. 0")
total_label.pack()

root.mainloop()
