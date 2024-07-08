import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from pdf2docx import Converter
from PIL import Image
import fitz  # PyMuPDF
from PyPDF2 import PdfFileReader, PdfFileWriter
import os

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OGConverter1.0")
        self.root.geometry("800x600")

        self.create_widgets()

    def create_widgets(self):
        self.label = tk.Label(self.root, text="Lütfen dönüştürmek veya şifrelemek istediğiniz dosyayı seçin:")
        self.label.pack(pady=10)

        self.file_entry = tk.Entry(self.root, width=70)
        self.file_entry.pack(pady=5)

        self.browse_button = tk.Button(self.root, text="Gözat", command=self.browse_file)
        self.browse_button.pack(pady=5)

        self.convert_word_to_pdf_button = tk.Button(self.root, text="Word'den PDF'e Çevir", command=self.convert_word_to_pdf)
        self.convert_word_to_pdf_button.pack(pady=10)

        self.convert_pdf_to_word_button = tk.Button(self.root, text="PDF'ten Word'e Çevir", command=self.convert_pdf_to_word)
        self.convert_pdf_to_word_button.pack(pady=10)

        self.convert_jpg_to_pdf_button = tk.Button(self.root, text="JPG'den PDF'e Çevir", command=self.convert_jpg_to_pdf)
        self.convert_jpg_to_pdf_button.pack(pady=10)

        self.convert_pdf_to_jpg_button = tk.Button(self.root, text="PDF'ten JPG'ye Çevir", command=self.convert_pdf_to_jpg)
        self.convert_pdf_to_jpg_button.pack(pady=10)

        self.encrypt_pdf_button = tk.Button(self.root, text="PDF'i Şifrele", command=self.encrypt_pdf)
        self.encrypt_pdf_button.pack(pady=10)

        self.result_text = tk.Text(self.root, height=15, width=90)
        self.result_text.pack(pady=20)

    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def convert_word_to_pdf(self):
        file_path = self.file_entry.get()
        if not file_path.endswith('.docx'):
            messagebox.showerror("Hata", "Lütfen bir Word dosyası seçin.")
            return
        try:
            pdf_path = file_path.replace('.docx', '.pdf')
            doc = Document(file_path)
            doc.save(pdf_path)
            self.result_text.insert(tk.END, f"Başarıyla dönüştürüldü: {pdf_path}\n")
        except Exception as e:
            self.result_text.insert(tk.END, f"Dönüştürme hatası: {str(e)}\n")

    def convert_pdf_to_word(self):
        file_path = self.file_entry.get()
        if not file_path.endswith('.pdf'):
            messagebox.showerror("Hata", "Lütfen bir PDF dosyası seçin.")
            return
        try:
            word_path = file_path.replace('.pdf', '.docx')
            cv = Converter(file_path)
            cv.convert(word_path, start=0, end=None)
            cv.close()
            self.result_text.insert(tk.END, f"Başarıyla dönüştürüldü: {word_path}\n")
        except Exception as e:
            self.result_text.insert(tk.END, f"Dönüştürme hatası: {str(e)}\n")

    def convert_jpg_to_pdf(self):
        file_path = self.file_entry.get()
        if not file_path.lower().endswith('.jpg') and not file_path.lower().endswith('.jpeg'):
            messagebox.showerror("Hata", "Lütfen bir JPG dosyası seçin.")
            return
        try:
            pdf_path = file_path.replace('.jpg', '.pdf').replace('.jpeg', '.pdf')
            image = Image.open(file_path)
            image.convert('RGB').save(pdf_path)
            self.result_text.insert(tk.END, f"Başarıyla dönüştürüldü: {pdf_path}\n")
        except Exception as e:
            self.result_text.insert(tk.END, f"Dönüştürme hatası: {str(e)}\n")

    def convert_pdf_to_jpg(self):
        file_path = self.file_entry.get()
        if not file_path.endswith('.pdf'):
            messagebox.showerror("Hata", "Lütfen bir PDF dosyası seçin.")
            return
        try:
            pdf_document = fitz.open(file_path)
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                pix = page.get_pixmap()
                jpg_path = file_path.replace('.pdf', f'_{page_num+1}.jpg')
                pix.save(jpg_path)
                self.result_text.insert(tk.END, f"Başarıyla dönüştürüldü: {jpg_path}\n")
        except Exception as e:
            self.result_text.insert(tk.END, f"Dönüştürme hatası: {str(e)}\n")

    def encrypt_pdf(self):
        file_path = self.file_entry.get()
        if not file_path.endswith('.pdf'):
            messagebox.showerror("Hata", "Lütfen bir PDF dosyası seçin.")
            return
        password = filedialog.askstring("Şifre", "Lütfen PDF dosyasını şifrelemek için bir parola girin:")
        if not password:
            messagebox.showerror("Hata", "Parola boş olamaz.")
            return
        try:
            pdf_reader = PdfFileReader(file_path)
            pdf_writer = PdfFileWriter()
            for page_num in range(pdf_reader.getNumPages()):
                pdf_writer.addPage(pdf_reader.getPage(page_num))
            pdf_writer.encrypt(user_pwd=password, owner_pwd=None, use_128bit=True)
            encrypted_pdf_path = file_path.replace('.pdf', '_encrypted.pdf')
            with open(encrypted_pdf_path, 'wb') as f:
                pdf_writer.write(f)
            self.result_text.insert(tk.END, f"Başarıyla şifrelendi: {encrypted_pdf_path}\n")
        except Exception as e:
            self.result_text.insert(tk.END, f"Şifreleme hatası: {str(e)}\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()
