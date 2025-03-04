import os
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from pdf2docx import Converter
from pdf2image import convert_from_path
import pytesseract
import customtkinter as ctk
from tkinter import filedialog, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, mm
from docx import Document
from docx.shared import Inches
from PIL import Image

# Ukuran kertas F4
F4 = (210 * mm, 330 * mm)

# Fungsi untuk menggabungkan PDF
def merge_pdfs():
    file_paths = filedialog.askopenfilenames(title="Pilih File PDF", filetypes=[("PDF Files", "*.pdf")])
    if file_paths:
        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_pdf:
            try:
                merger = PdfMerger()
                for path in file_paths:
                    merger.append(path)
                merger.write(output_pdf)
                merger.close()
                messagebox.showinfo("Sukses", f"File PDF berhasil digabung: {output_pdf}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal menggabungkan PDF: {str(e)}")

# Fungsi untuk konversi PDF ke Word
def pdf_to_word():
    file_path = filedialog.askopenfilename(title="Pilih File PDF", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        output_docx = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Files", "*.docx")])
        if output_docx:
            try:
                cv = Converter(file_path)
                cv.convert(output_docx)
                cv.close()
                messagebox.showinfo("Sukses", f"File Word berhasil dibuat: {output_docx}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal mengonversi PDF: {str(e)}")

# Fungsi untuk konversi file ke PDF
def convert_to_pdf():
    file_path = filedialog.askopenfilename(title="Pilih File", filetypes=[("Text Files", "*.txt"), ("Image Files", "*.jpg *.png")])
    if file_path:
        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_pdf:
            try:
                if file_path.endswith(".txt"):
                    with open(file_path, "r", encoding="utf-8") as file:
                        text = file.read()
                    c = canvas.Canvas(output_pdf, pagesize=A4)
                    c.drawString(100, 750, text)
                    c.save()
                else:
                    img = Image.open(file_path)
                    img.save(output_pdf, "PDF", resolution=100.0)
                messagebox.showinfo("Sukses", f"File PDF berhasil dibuat: {output_pdf}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal mengonversi file: {str(e)}")

# Fungsi utama GUI
def main():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("PDF Tools")
    root.geometry("500x400")

    title_label = ctk.CTkLabel(root, text="PDF Tools", font=("Arial", 24, "bold"))
    title_label.pack(pady=20)

    buttons = [
        ("Gabung PDF", merge_pdfs),
        ("PDF ke Word", pdf_to_word),
        ("Konversi ke PDF", convert_to_pdf)
    ]

    for text, command in buttons:
        btn = ctk.CTkButton(root, text=text, command=command, width=250, height=40)
        btn.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
