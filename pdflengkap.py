import os
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from pdf2docx import Converter
from pdf2image import convert_from_path
import pytesseract
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, mm
from docx import Document
from docx.shared import Inches
from PIL import Image

# Ukuran kertas F4
F4 = (210 * mm, 330 * mm)  # 210mm x 330mm

# Fungsi untuk gabung PDF
def merge_pdfs():
    file_paths = filedialog.askopenfilenames(title="Pilih File PDF untuk Digabung", filetypes=[("PDF Files", "*.pdf")])
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

# Fungsi untuk PDF to Word Biasa
def pdf_to_word_normal():
    file_path = filedialog.askopenfilename(title="Pilih File PDF untuk Convert ke Word", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        output_docx = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Files", "*.docx")])
        if output_docx:
            try:
                cv = Converter(file_path)
                cv.convert(output_docx)
                cv.close()
                set_page_size(output_docx, page_size_var.get())  # Atur ukuran halaman
                messagebox.showinfo("Sukses", f"File Word berhasil dibuat: {output_docx}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal convert PDF: {str(e)}")

# Fungsi untuk PDF to Word Gambar (OCR)
def pdf_to_word_image():
    file_path = filedialog.askopenfilename(title="Pilih File PDF untuk Convert ke Word", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        output_docx = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Files", "*.docx")])
        if output_docx:
            try:
                images = convert_from_path(file_path)
                text = ""
                for image in images:
                    text += pytesseract.image_to_string(image, lang='ind')  # 'ind' untuk bahasa Indonesia
                with open(output_docx, "w", encoding="utf-8") as file:
                    file.write(text)
                set_page_size(output_docx, page_size_var.get())  # Atur ukuran halaman
                messagebox.showinfo("Sukses", f"File Word berhasil dibuat dengan OCR: {output_docx}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal convert PDF: {str(e)}")

# Fungsi untuk PDF to Word Maximal (Gabungan)
def pdf_to_word_maximal():
    file_path = filedialog.askopenfilename(title="Pilih File PDF untuk Convert ke Word", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        output_docx = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Files", "*.docx")])
        if output_docx:
            try:
                # Coba konversi dengan pdf2docx
                cv = Converter(file_path)
                cv.convert(output_docx)
                cv.close()
                set_page_size(output_docx, page_size_var.get())  # Atur ukuran halaman
                messagebox.showinfo("Sukses", f"File Word berhasil dibuat: {output_docx}")
            except:
                # Jika gagal, gunakan OCR
                try:
                    images = convert_from_path(file_path)
                    text = ""
                    for image in images:
                        text += pytesseract.image_to_string(image, lang='ind')
                    with open(output_docx, "w", encoding="utf-8") as file:
                        file.write(text)
                    set_page_size(output_docx, page_size_var.get())  # Atur ukuran halaman
                    messagebox.showinfo("Sukses", f"File Word berhasil dibuat dengan OCR: {output_docx}")
                except Exception as e:
                    messagebox.showerror("Error", f"Gagal convert PDF: {str(e)}")

# Fungsi untuk mengatur ukuran halaman di Word
def set_page_size(docx_path, page_size):
    doc = Document(docx_path)
    section = doc.sections[0]
    if page_size == "A4":
        section.page_width = Inches(8.27)  # Lebar A4: 210 mm
        section.page_height = Inches(11.69)  # Tinggi A4: 297 mm
    elif page_size == "F4":
        section.page_width = Inches(8.27)  # Lebar F4: 210 mm
        section.page_height = Inches(13.00)  # Tinggi F4: 330 mm
    doc.save(docx_path)

# Fungsi untuk convert file ke PDF
def convert_to_pdf():
    file_path = filedialog.askopenfilename(title="Pilih File untuk Convert ke PDF", filetypes=[("Text Files", "*.txt"), ("Image Files", "*.jpg *.png")])
    if file_path:
        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_pdf:
            try:
                # Pilih ukuran halaman
                page_size = page_size_var.get()  # Ambil nilai dari dropdown
                if page_size == "A4":
                    size = A4
                elif page_size == "F4":
                    size = F4

                if file_path.endswith(".txt"):
                    # Convert teks ke PDF
                    with open(file_path, "r", encoding="utf-8") as file:
                        text = file.read()
                    c = canvas.Canvas(output_pdf, pagesize=size)
                    c.drawString(100, 750, "Isi File Teks:")
                    c.drawString(100, 730, text)
                    c.save()
                elif file_path.endswith((".jpg", ".png")):
                    # Convert gambar ke PDF
                    img = Image.open(file_path)
                    img.save(output_pdf, "PDF", resolution=100.0)
                messagebox.showinfo("Sukses", f"File PDF berhasil dibuat: {output_pdf}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal convert file: {str(e)}")

# Fungsi untuk kompres PDF
def compress_pdf():
    file_path = filedialog.askopenfilename(title="Pilih File PDF untuk Kompres", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_pdf:
            try:
                # Buka file PDF
                reader = PdfReader(file_path)
                writer = PdfWriter()

                # Tambahkan halaman ke writer
                for page in reader.pages:
                    writer.add_page(page)

                # Kompres file PDF
                writer.add_metadata(reader.metadata)
                with open(output_pdf, "wb") as f:
                    writer.write(f)

                # Cek ukuran file
                file_size = os.path.getsize(output_pdf) / 1024  # Ukuran dalam KB
                if file_size > 1024:  # Jika lebih dari 1 MB
                    messagebox.showwarning("Peringatan", f"Ukuran file masih {file_size:.2f} KB. Coba kompres lagi.")
                else:
                    messagebox.showinfo("Sukses", f"File PDF berhasil dikompres: {output_pdf}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal kompres PDF: {str(e)}")

# Fungsi untuk pisahkan PDF
def split_pdf():
    file_path = filedialog.askopenfilename(title="Pilih File PDF untuk Dipisahkan", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        try:
            # Buka file PDF
            reader = PdfReader(file_path)
            total_pages = len(reader.pages)

            # Minta pengguna memilih halaman
            selected_pages = simpledialog.askstring("Pilih Halaman", f"Masukkan nomor halaman (contoh: 1,3,5 atau 1-5):")
            if selected_pages:
                output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
                if output_pdf:
                    writer = PdfWriter()
                    if "-" in selected_pages:
                        # Jika format range (contoh: 1-5)
                        start, end = map(int, selected_pages.split("-"))
                        for i in range(start - 1, end):
                            writer.add_page(reader.pages[i])
                    else:
                        # Jika format list (contoh: 1,3,5)
                        pages = list(map(int, selected_pages.split(",")))
                        for page in pages:
                            writer.add_page(reader.pages[page - 1])

                    # Simpan file PDF
                    with open(output_pdf, "wb") as f:
                        writer.write(f)
                    messagebox.showinfo("Sukses", f"File PDF berhasil dipisahkan: {output_pdf}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal memisahkan PDF: {str(e)}")

# GUI Aplikasi
def main():
    global page_size_var  # Variabel global

    # Set tema dan warna
    ctk.set_appearance_mode("System")  # Tema: System, Light, Dark
    ctk.set_default_color_theme("blue")  # Warna: blue, green, dark-blue

    # Buat window
    root = ctk.CTk()
    root.title("PDF Tools by PeekCode (@masulin00)")
    root.geometry("800x600")

    # Judul Aplikasi
    title_label = ctk.CTkLabel(root, text="PDF Tools", font=("Arial", 24, "bold"))
    title_label.pack(pady=20)

    # Frame utama (layout kiri dan kanan)
    main_frame = ctk.CTkFrame(root)
    main_frame.pack(pady=10, padx=20, fill="both", expand=True)

    # Kolom Kiri (Fitur Utama)
    left_frame = ctk.CTkFrame(main_frame)
    left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

    # Frame untuk Gabung PDF
    merge_frame = ctk.CTkFrame(left_frame)
    merge_frame.pack(pady=10, padx=10, fill="x")

    merge_label = ctk.CTkLabel(merge_frame, text="Gabung PDF", font=("Arial", 16))
    merge_label.pack(pady=5)

    merge_button = ctk.CTkButton(merge_frame, text="Gabung PDF", command=merge_pdfs, width=200, height=40)
    merge_button.pack(pady=10)

    # Frame untuk Convert PDF to Word
    convert_frame = ctk.CTkFrame(left_frame)
    convert_frame.pack(pady=10, padx=10, fill="x")

    convert_label = ctk.CTkLabel(convert_frame, text="Convert PDF to Word", font=("Arial", 16))
    convert_label.pack(pady=5)

    normal_button = ctk.CTkButton(convert_frame, text="PDF to Word Biasa", command=pdf_to_word_normal, width=200, height=40)
    normal_button.pack(pady=5)

    image_button = ctk.CTkButton(convert_frame, text="PDF to Word Gambar", command=pdf_to_word_image, width=200, height=40)
    image_button.pack(pady=5)

    maximal_button = ctk.CTkButton(convert_frame, text="PDF to Word Maximal", command=pdf_to_word_maximal, width=200, height=40)
    maximal_button.pack(pady=5)

    # Frame untuk Convert File ke PDF
    convert_pdf_frame = ctk.CTkFrame(left_frame)
    convert_pdf_frame.pack(pady=10, padx=10, fill="x")

    convert_pdf_label = ctk.CTkLabel(convert_pdf_frame, text="Convert File ke PDF", font=("Arial", 16))
    convert_pdf_label.pack(pady=5)

    # Dropdown untuk memilih ukuran halaman
    page_size_var = ctk.StringVar(value="A4")  # Default: A4
    page_size_menu = ctk.CTkOptionMenu(convert_pdf_frame, variable=page_size_var, values=["A4", "F4"])
    page_size_menu.pack(pady=5)

    convert_pdf_button = ctk.CTkButton(convert_pdf_frame, text="Convert ke PDF", command=convert_to_pdf, width=200, height=40)
    convert_pdf_button.pack(pady=10)

    # Kolom Kanan (Fitur Tambahan)
    right_frame = ctk.CTkFrame(main_frame)
    right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

    # Frame untuk Kompres PDF
    compress_frame = ctk.CTkFrame(right_frame)
    compress_frame.pack(pady=10, padx=10, fill="x")

    compress_label = ctk.CTkLabel(compress_frame, text="Kompres PDF", font=("Arial", 16))
    compress_label.pack(pady=5)

    compress_button = ctk.CTkButton(compress_frame, text="Kompres PDF (<1 MB)", command=compress_pdf, width=200, height=40)
    compress_button.pack(pady=10)

    # Frame untuk Pisahkan PDF
    split_frame = ctk.CTkFrame(right_frame)
    split_frame.pack(pady=10, padx=10, fill="x")

    split_label = ctk.CTkLabel(split_frame, text="Pisahkan PDF", font=("Arial", 16))
    split_label.pack(pady=5)

    split_button = ctk.CTkButton(split_frame, text="Pisahkan PDF", command=split_pdf, width=200, height=40)
    split_button.pack(pady=10)

    # Jalankan aplikasi
    root.mainloop()

if __name__ == "__main__":
    main()