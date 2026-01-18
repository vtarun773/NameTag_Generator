import customtkinter as ctk
from tkinter import filedialog, messagebox
from ppt_generator import generate_ppt
import os

# ---------------- Functions ----------------
def browse_csv():
    path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if path: csv_var.set(path)

def browse_template():
    path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if path: template_var.set(path)

def browse_images():
    path = filedialog.askdirectory()
    if path: images_var.set(path)

def browse_output():
    folder = filedialog.askdirectory()
    if folder: output_var.set(os.path.join(folder, "output.pptx"))

def run_generator():
    if not csv_var.get(): return messagebox.showerror("Error", "Please select CSV file")
    if not template_var.get(): return messagebox.showerror("Error", "Please select template PPT")
    if not images_var.get(): return messagebox.showerror("Error", "Please select images folder")
    if not output_var.get(): return messagebox.showerror("Error", "Please select output folder")

    email_suffix = email_suffix_var.get().strip()
    if not email_suffix: return messagebox.showerror("Error", "Email suffix cannot be empty")

    try:
        generate_ppt(csv_var.get(), template_var.get(), images_var.get(), output_var.get(), email_suffix)
        messagebox.showinfo("Success", "PowerPoint generated successfully!\n\nFile saved as:\noutput.pptx")
    except Exception as e:
        messagebox.showerror("Failed", str(e))

# ---------------- UI Setup ----------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Employee Nametag Generator")
root.geometry("720x450")
root.resizable(True, True)

# ---------------- Variables ----------------
csv_var = ctk.StringVar()
template_var = ctk.StringVar()
images_var = ctk.StringVar()
output_var = ctk.StringVar()
email_suffix_var = ctk.StringVar(value="@samsung.com")

# ---------------- Center Frame ----------------
center_frame = ctk.CTkFrame(root, fg_color="transparent")  # transparent so dark background shows
center_frame.grid(row=0, column=0, sticky="nsew")

# Configure root grid to center frame
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

pad_y = 12

# ---------------- Layout using grid inside center frame ----------------
center_frame.grid_columnconfigure(0, weight=1)
center_frame.grid_columnconfigure(1, weight=2)
center_frame.grid_columnconfigure(2, weight=1)

def create_row(label_text, var, browse_cmd, row_index):
    ctk.CTkLabel(center_frame, text=label_text, font=("Segoe UI", 14)).grid(row=row_index, column=0, sticky="e", padx=10, pady=pad_y)
    entry = ctk.CTkEntry(center_frame, textvariable=var, width=400, height=35, corner_radius=20)
    entry.grid(row=row_index, column=1, sticky="ew", padx=10)
    if browse_cmd:
        ctk.CTkButton(center_frame, text="Browse", command=browse_cmd, width=100, height=35, corner_radius=20, font=("Segoe UI", 14)).grid(row=row_index, column=2, padx=10)

create_row("CSV File", csv_var, browse_csv, 0)
create_row("Template PPT", template_var, browse_template, 1)
create_row("Images Folder", images_var, browse_images, 2)
create_row("Output Folder", output_var, browse_output, 3)
create_row("Email Suffix", email_suffix_var, None, 4)

# Generate button
ctk.CTkButton(center_frame, text="Generate Nametags", command=run_generator,
              width=200, height=50, corner_radius=25, font=("Segoe UI", 14, "bold")).grid(row=5, column=0, columnspan=3, pady=30)

# Bottom-right "made by" label
made_by_label = ctk.CTkLabel(root, text="developed by tarun1.verma", font=("Segoe UI", 10))
made_by_label.place(relx=1.0, rely=1.0, x=-10, y=-5, anchor="se")

# ---------------- Make the frame expand vertically ----------------
for i in range(6):
    center_frame.grid_rowconfigure(i, weight=1)

root.mainloop()
