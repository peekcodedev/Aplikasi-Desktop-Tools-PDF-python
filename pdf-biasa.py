import os
from PyPDF2 import PdfMerger
from pdf2docx import Converter
from reportlab.pdfgen import canvas
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Fungsi untuk convert file ke PDF
def convert_to_pdf():
    file_path = filedialog.askopenfilename(title="Pilih File untuk Convert ke PDF", filetypes=[("All Files", "*.*")])
    if file_path:
        output_pdf = os.path.splitext(file_path)[0] + ".pdf"
        c = canvas.Canvas(output_pdf)
        c.drawString(100, 750, "Ini adalah file PDF yang dibuat dari Python!")
        c.save()
        messagebox.showinfo("Sukses", f"File PDF berhasil dibuat: {output_pdf}")

# Fungsi untuk gabung PDF
def merge_pdfs():
    file_paths = filedialog.askopenfilenames(title="Pilih File PDF untuk Digabung", filetypes=[("PDF Files", "*.pdf")])
    if file_paths:
        merger = PdfMerger()
        for path in file_paths:
            merger.append(path)
        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_pdf:
            merger.write(output_pdf)
            merger.close()
            messagebox.showinfo("Sukses", f"File PDF berhasil digabung: {output_pdf}")

# Fungsi untuk convert PDF ke Word
def convert_to_word():
    file_path = filedialog.askopenfilename(title="Pilih File PDF untuk Convert ke Word", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        output_docx = os.path.splitext(file_path)[0] + ".docx"
        cv = Converter(file_path)
        cv.convert(output_docx)
        cv.close()
        messagebox.showinfo("Sukses", f"File Word berhasil dibuat: {output_docx}")

# GUI Aplikasi dengan CustomTkinter
def main():
    # Set tema dan warna
    ctk.set_appearance_mode("System")  # Tema: System, Light, Dark
    ctk.set_default_color_theme("blue")  # Warna: blue, green, dark-blue

    # Buat window
    root = ctk.CTk()
    root.title("PDF Tools by PeekCode")
    root.geometry("500x400")

    # Judul Aplikasi
    title_label = ctk.CTkLabel(root, text="PDF Tools", font=("Arial", 24, "bold"))
    title_label.pack(pady=20)

    # Tombol Convert File ke PDF
    convert_pdf_button = ctk.CTkButton(root, text="Convert File ke PDF", command=convert_to_pdf, width=200, height=40)
    convert_pdf_button.pack(pady=10)

    # Tombol Gabung PDF
    merge_pdf_button = ctk.CTkButton(root, text="Gabung PDF", command=merge_pdfs, width=200, height=40)
    merge_pdf_button.pack(pady=10)

    # Tombol Convert PDF ke Word
    convert_word_button = ctk.CTkButton(root, text="Convert PDF ke Word", command=convert_to_word, width=200, height=40)
    convert_word_button.pack(pady=10)

    # Jalankan aplikasi
    root.mainloop()

if __name__ == "__main__":
    main()