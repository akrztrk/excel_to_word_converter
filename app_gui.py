import tkinter as tk
from tkinter import filedialog, messagebox
from converter.controller import ConverterController

class ExcelToWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel'den Word'e Dönüştürücü")

        self.selected_file = None
        self.output_folder = None

        # Excel dosyası seçme bölümü
        self.label_file = tk.Label(root, text="Bir Excel dosyası seçin:")
        self.label_file.pack(pady=5)

        self.select_file_button = tk.Button(root, text="Dosya Seç", command=self.select_file)
        self.select_file_button.pack(pady=5)

        # Çıktı klasörü seçme bölümü
        self.label_folder = tk.Label(root, text="Çıktı klasörü seçin:")
        self.label_folder.pack(pady=5)

        self.select_folder_button = tk.Button(root, text="Klasör Seç", command=self.select_folder)
        self.select_folder_button.pack(pady=5)

        self.selected_folder_label = tk.Label(root, text="Henüz klasör seçilmedi")
        self.selected_folder_label.pack(pady=5)

        # Dönüştür butonu
        self.convert_button = tk.Button(root, text="Dönüştür", command=self.convert_file, state=tk.DISABLED)
        self.convert_button.pack(pady=10)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls")]
        )
        if file_path:
            self.selected_file = file_path
            self.label_file.config(text=f"Seçilen dosya:\n{file_path}")
            self.update_convert_button_state()

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder = folder_path
            self.selected_folder_label.config(text=f"Seçilen klasör:\n{folder_path}")
            self.update_convert_button_state()

    def update_convert_button_state(self):
        # Hem dosya hem de çıktı klasörü seçildiyse aktif et
        if self.selected_file and self.output_folder:
            self.convert_button.config(state=tk.NORMAL)

    def convert_file(self):
        if not self.selected_file:
            messagebox.showerror("Hata", "Lütfen önce bir Excel dosyası seçin.")
            return
        if not self.output_folder:
            messagebox.showerror("Hata", "Lütfen önce çıktı klasörünü seçin.")
            return

        try:
            controller = ConverterController(self.selected_file, self.output_folder)
            controller.convert_all_sheets()
            messagebox.showinfo("Başarılı", f"Tüm sayfalar '{self.output_folder}' klasörüne dönüştürüldü!")
        except Exception as e:
            messagebox.showerror("Hata", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToWordApp(root)
    root.mainloop()

