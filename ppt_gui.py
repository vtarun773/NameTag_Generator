import tkinter as tk
from tkinter import filedialog, messagebox
from ppt_generator import generate_ppt
import os


def browse_csv():
    path = filedialog.askopenfilename(
        filetypes=[("CSV files", "*.csv")]
    )
    if path:
        csv_var.set(path)


def browse_template():
    path = filedialog.askopenfilename(
        filetypes=[("PowerPoint files", "*.pptx")]
    )
    if path:
        template_var.set(path)


def browse_images():
    path = filedialog.askdirectory()
    if path:
        images_var.set(path)


def browse_output():
    folder = filedialog.askdirectory()
    if folder:
        output_var.set(os.path.join(folder, "output.pptx"))


def run_generator():
    if not csv_var.get():
        messagebox.showerror("Error", "Please select CSV file")
        return
    if not template_var.get():
        messagebox.showerror("Error", "Please select template PPT")
        return
    if not images_var.get():
        messagebox.showerror("Error", "Please select images folder")
        return
    if not output_var.get():
        messagebox.showerror("Error", "Please select output folder")
        return

    email_suffix = email_suffix_var.get().strip()
    if not email_suffix:
        messagebox.showerror("Error", "Email suffix cannot be empty")
        return

    try:
        generate_ppt(
            csv_var.get(),
            template_var.get(),
            images_var.get(),
            output_var.get(),
            email_suffix
        )

        messagebox.showinfo(
            "Success",
            "PowerPoint generated successfully!\n\nFile saved as:\noutput.pptx"
        )

        root.destroy()  # Exit after OK

    except Exception as e:
        messagebox.showerror("Failed", str(e))


# ---------------- UI ----------------
root = tk.Tk()
root.title("Employee Nametag Generator")
root.geometry("700x380")
root.resizable(True, True)

csv_var = tk.StringVar()
template_var = tk.StringVar()
images_var = tk.StringVar()
output_var = tk.StringVar()

# Default suffix
email_suffix_var = tk.StringVar(value="@samsung.com")


def row(label, var, browse_cmd, r, readonly=True):
    tk.Label(
        root,
        text=label,
        anchor="w",
        font=("Segoe UI", 9)
    ).grid(row=r, column=0, padx=10, pady=10, sticky="w")

    tk.Entry(
        root,
        textvariable=var,
        width=48,
        state="readonly" if readonly else "normal"
    ).grid(row=r, column=1)

    if browse_cmd:
        tk.Button(
            root,
            text="Browse",
            command=browse_cmd,
            width=10
        ).grid(row=r, column=2, padx=5)


row("CSV File", csv_var, browse_csv, 0)
row("Template PPT", template_var, browse_template, 1)
row("Images Folder", images_var, browse_images, 2)
row("Output Folder", output_var, browse_output, 3)
row("Email Suffix", email_suffix_var, None, 4, readonly=False)

tk.Button(
    root,
    text="Generate Nametags",
    command=run_generator,
    bg="#0078D7",
    fg="white",
    font=("Segoe UI", 10, "bold"),
    width=22
).grid(row=6, column=1, pady=30)

root.mainloop()
