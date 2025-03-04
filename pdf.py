import os
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from pdf2docx import Converter
from reportlab.pdfgen import canvas
from tkinter import Tk, Button, Label, filedialog, messagebox, Entry, StringVar

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

# GUI Aplikasi
def main():
    root = Tk()
    root.title("PDF Tools by PEEKCODE")
    root.geometry("400x300")

    Label(root, text="Pilih Fungsi:").pack(pady=10)

    Button(root, text="Convert File ke PDF", command=convert_to_pdf).pack(pady=5)
    Button(root, text="Gabung PDF", command=merge_pdfs).pack(pady=5)
    Button(root, text="Convert PDF ke Word", command=convert_to_word).pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()