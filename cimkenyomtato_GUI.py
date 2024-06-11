import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import time
import subprocess
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to create and save the label
def print_label(text, qr_text, copies, filename):
    label_width = 30*3
    label_height = 15*3
    custom_font = ImageFont.truetype("arial.ttf", 100)

    # Create label image
    label_image = Image.new('RGB', (int(label_width * 10), int(label_height * 10)), color='white')
    draw = ImageDraw.Draw(label_image)
    
    # Add QR code to the left side
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=11.0, border=4)
    qr.add_data(qr_text)
    qr.make(fit=True)
    qr_image = qr.make_image(fill_color="black", back_color="white")
    label_image.paste(qr_image, (4, 4))
    
    # Add text to the right side, vertically centered
    text_width, text_height = draw.textsize(text, font=custom_font)
    text_start_x = ((label_width * 10 - text_width) / 2) +120
    text_start_y = ((label_height * 10 - text_height) / 2) - 60
    draw.text((text_start_x, text_start_y), text, fill='black', font=custom_font)
    
    # Save label to file
    label_image.save(filename)

def start_printing():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            messagebox.showerror("Error", "No file selected!")
            return

        df = pd.read_excel(file_path)
        df = df.sort_values(by=df.columns[3])
        
        file_list = []

        for index, row in df.iterrows():
            id_string = str(row.iloc[0])
            if id_string[-1].isalpha():
                last_characters = id_string[-5:]
            else:
                last_characters = id_string[-4:]
            label_text = "ZE" + last_characters
            qr_text = f"{row.iloc[0]}\n{row.iloc[1]}"
            copies = int(row.iloc[2])
            if copies != 0:
                if copies > 9:
                    copies = 1
                for i in range(copies):
                    filename = f"C:/Users/KAZ5EGR/Desktop/cimke/label_{row.iloc[0]}_{i}.png"
                    print_label(label_text, qr_text, copies, filename)
                    file_list.append(filename)
                    time.sleep(0.2)
        
        messagebox.showinfo("Success", "Labels created and saved successfully!")
        run_print_script(file_list)
    except Exception as e:
        messagebox.showerror("Error", str(e))

def run_print_script(file_list):
    script_path = r'C:\Users\KAZ5EGR\Desktop\cimke\nyomtatas.ps1'
    result = subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-File', script_path], capture_output=True, text=True)
    if result.stdout:
        print("STDOUT:", result.stdout)
    if result.stderr:
        print("STDERR:", result.stderr)
    if result.returncode == 0:
        print("Print command executed successfully.")
        delete_files(file_list)
    else:
        print(f"Print command failed with return code: {result.returncode}")

def delete_files(file_list):
    for file in file_list:
        try:
            os.remove(file)
            print(f"Deleted file: {file}")
        except OSError as e:
            print(f"Error deleting file {file}: {e.strerror}")

# Tkinter UI setup
root = tk.Tk()
root.title("Label Printer")

# Create and place the button to select file and start printing
button_print = tk.Button(root, text="Select Excel File and Print", command=start_printing)
button_print.pack(pady=20)

root.mainloop()
