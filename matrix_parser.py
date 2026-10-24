import re
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import pyperclip


class MatrixGUIApp:
    """Havacılık Lisans Matris Verilerini Panodan Okuyan,
    'Sınav' veriliş şekillerini de tanıyan Masaüstü Uygulaması.
    """

    def __init__(self, root: tk.Tk, db_path: str = "havacilik_egitim.db"):
        self.root = root
        self.db_path = db_path
        self.table_name = "lisans_ders_matrisi"

        self.root.title("Havacılık Lisans & Ders Matrisi Aktarıcı")
        self.root.geometry("1220x680")
        self.root.minsize(1000, 500)

        self._create_widgets()
        self.load_all_data_from_db()

    def _create_widgets(self):
        # 1. İşlem Paneli ve Butonlar
        top_frame = ttk.LabelFrame(self.root, text=" İşlem Paneli ", padding=10)
        top_frame.pack(fill="x", padx=15, pady=10)

        btn_clipboard = ttk.Button(
            top_frame, text="📋 Panodan Oku ve Kaydet", command=self.process_clipboard
        )
        btn_clipboard.pack(side="left", padx=5)

        btn_delete_selected = ttk.Button(
            top_frame, text="🗑️ Seçilenleri Sil (Çoklu)", command=self.delete_selected_records
        )
        btn_delete_selected.pack(side="left", padx=5)

        btn_export_excel = ttk.Button(
            top_frame, text="📊 Excel'e Aktar", command=self.export_to_excel
        )
        btn_export_excel.pack(side="left", padx=5)

        btn_export_csv = ttk.Button(
            top_frame, text="📄 CSV'ye Aktar", command=self.export_to_csv
        )
        btn_export_csv.pack(side="left", padx=5)

        btn_refresh = ttk.Button(
            top_frame, text="🔄 Tümünü Göster", command=self.load_all_data_from_db
        )
        btn_refresh.pack(side="left", padx=5)

        btn_exit = ttk.Button(top_frame, text="❌ Çıkış", command=self.root.quit)
        btn_exit.pack(side="right", padx=5)

        # 2. Durum Bilgisi
        self.lbl_status = ttk.Label(
            self.root, text="Hazır. Panodan veri aktarmak için butona basın.", font=("Arial", 10, "italic")
        )
        self.lbl_status.pack(anchor="w", padx=20, pady=2)

        # 3. Tablo (Treeview)
        table_frame = ttk.Frame(self.root, padding=10)
        table_frame.pack(fill="both", expand=True, padx=15, pady=5)

        columns = (
            "lisans_adi", "ders_kodu_ve_adi", "egitim_cinsi", 
            "verilis_sekli", "sinav_tipi", "gecme_notu", 
            "zorunluluk", "sure", "yenileme_suresi"
        )

        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")
        
        headers = {
            "lisans_adi": ("Lisans Adı", 220),
            "ders_kodu_ve_adi": ("Ders Kodu ve Adı", 320),
            "egitim_cinsi": ("Eğitim Cinsi", 100),
            "verilis_sekli": ("Veriliş Şekli", 110),
            "sinav_tipi": ("Sınav Tipi", 150),
            "gecme_notu": ("Geçme Notu", 80),
            "zorunluluk": ("Zorunluluk", 80),
            "sure": ("Süre", 70),
            "yenileme_suresi": ("Yenileme", 80)
        }

        for col_id, (col_name, col_width) in headers.items():
            self.tree.heading(col_id, text=col_name)
            self.tree.column(col_id, width=col_width, anchor="center" if col_width < 100 else "w")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

    def get_clipboard_text(self) -> str:
        try:
            return pyperclip.paste()
        except Exception:
            try:
                return self.root.clipboard_get()
            except Exception:
                return ""

    def clean_pre_matris(self, raw_text: str) -> str:
        """Sadece 'Matris' kelimesini bulup öncesini keser."""
        match = re.search(r"Matris", raw_text, re.IGNORECASE)
        if match:
            return raw_text[match.start():]
        return None

    def extract_lisans_adi(self, cleaned_text: str) -> str:
        """Matris kelimesinden sonra gelen lisans adını keser."""
        lines = [line.strip() for line in cleaned_text.splitlines() if line.strip()]
        if not lines:
            return None

        first_line_rest = re.sub(r"^Matris\s*", "", lines[0], flags=re.IGNORECASE).strip()
        if first_line_rest:
            return first_line_rest
        
        return lines[1] if len(lines) > 1 else "Bilinmeyen Lisans"

    def parse_raw_text(self, raw_text: str):
        """Metni temizler ve belirlenen haritalama yapısına göre verileri ayıklar."""
        cleaned_text = self.clean_pre_matris(raw_text)
        if not cleaned_text:
            return None, []

        lisans_adi = self.extract_lisans_adi(cleaned_text)

        clean_stream = re.sub(r"\s+", " ", cleaned_text.replace("\xa0", " ").replace("\t", " ")).strip()

        # Ders başlangıçlarını ('U 01' vb.) indeks olarak bul
        starts = [m.start() for m in re.finditer(r"\bU\s*\d{2}\b", clean_stream)]
        if not starts:
            return lisans_adi, []

        blocks = []
        for i in range(len(starts)):
            start_idx = starts[i]
            end_idx = starts[i+1] if i + 1 < len(starts) else len(clean_stream)
            blocks.append(clean_stream[start_idx:end_idx].strip())

        # Ham Eğitim Cinsi -> DB verilis_sekli ('Sınav' eklendi)
        verilis_pattern = r"(Dijital\s+Ders|Dijital\s+Bilgi\s+Notu|Pratik\s+Ders|Pratik|Simülatör\s+Ders|Simülatör|Örgün\s+Ders|Örgün|Uzaktan\s+Eğitim|Sınav)"
        
        # Ham Veriliş Şekli + Sınav -> DB sinav_tipi
        sinav_pattern = r"((?:Gözetmenli\s+Online|Online|Bütünleşik-S|Bütünleşik|Uygulama)?\s*Sınav(?:lı|sız))"
        
        # Geçme Notu, Zorunluluk, Süre (Opsiyonel) ve Yenileme
        kuyruk_pattern = r"(\d+)\s+([\wÇĞİÖŞÜçğiöşü]+)\s+(?:(\d+\s*Saat)\s+)?(\d+\s*Yıl|∞|-)"

        anchor_regex = re.compile(
            rf"{verilis_pattern}\s+{sinav_pattern}\s+{kuyruk_pattern}",
            re.IGNORECASE
        )

        records = []
        for block in blocks:
            block_clean = re.sub(r"Genel\s*Ort:?\s*\d*.*$", "", block, flags=re.IGNORECASE).strip()

            match = anchor_regex.search(block_clean)
            if match:
                verilis_sekli = match.group(1).strip()
                sinav_tipi = match.group(2).strip()
                gecme_notu = int(match.group(3))
                zorunluluk = match.group(4).strip()
                sure = match.group(5).strip() if match.group(5) else "-"
                yenileme_suresi = match.group(6).strip()

                head_text = block_clean[:match.start()].strip()

                # Ham Eğitim Tipi -> DB egitim_cinsi
                egitim_cinsi_match = re.search(
                    r"(★?\s*(?:Uzmanlaşma|Temel|Başlangıç|-))\s*-?\s*$", 
                    head_text, 
                    re.IGNORECASE
                )

                if egitim_cinsi_match:
                    raw_egitim = egitim_cinsi_match.group(1)
                    egitim_cinsi = raw_egitim.replace("★", "").strip()
                    ders_kodu_ve_adi = head_text[:egitim_cinsi_match.start()].strip()
                else:
                    egitim_cinsi = "Belirtilmemiş"
                    ders_kodu_ve_adi = head_text

                ders_kodu_ve_adi = re.sub(r"^(?:Genel\s*Ort:?\s*\d*\s*)+", "", ders_kodu_ve_adi, flags=re.IGNORECASE).strip()

                records.append({
                    "lisans_adi": lisans_adi,
                    "ders_kodu_ve_adi": ders_kodu_ve_adi,
                    "egitim_cinsi": egitim_cinsi if egitim_cinsi != "-" else "Belirtilmemiş",
                    "verilis_sekli": verilis_sekli,
                    "sinav_tipi": sinav_tipi,
                    "gecme_notu": gecme_notu,
                    "zorunluluk": zorunluluk,
                    "sure": sure,
                    "yenileme_suresi": yenileme_suresi
                })

        return lisans_adi, records

    def process_clipboard(self):
        """Panodan okur ve kaydedilmeyi bekleyen verileri işler."""
        raw_text = self.get_clipboard_text()

        if not raw_text or not raw_text.strip():
            messagebox.showwarning("Uyarı", "Panoda kopyalanmış metin bulunamadı!")
            return

        lisans_adi, data = self.parse_raw_text(raw_text)

        if not lisans_adi:
            messagebox.showerror(
                "Format Hatası", 
                "Panodaki metinde 'Matris' kelimesi bulunamadı!\n\nLütfen 'Matris' ile başlayan alanı kopyaladığınızdan emin olun."
            )
            return

        if not data:
            messagebox.showerror("Format Hatası", f"'{lisans_adi}' altında geçerli ders tablosu (U 01 ...) bulunamadı.")
            return

        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (self.table_name,))
            table_exists = cursor.fetchone()

            existing_dersler = set()
            if table_exists:
                cursor.execute(
                    f"SELECT ders_kodu_ve_adi FROM {self.table_name} WHERE lisans_adi = ?", 
                    (lisans_adi,)
                )
                existing_dersler = {row[0] for row in cursor.fetchall()}

            new_records = [rec for rec in data if rec["ders_kodu_ve_adi"] not in existing_dersler]
            skipped_count = len(data) - len(new_records)

            if not new_records:
                self.lbl_status.config(
                    text=f"⚠️ '{lisans_adi}' altındaki kopyalanan tüm dersler ({skipped_count} ders) zaten veritabanında var!",
                    foreground="orange"
                )
                messagebox.showinfo("Mevcut Kayıtlar", f"'{lisans_adi}' lisansına ait bu derslerin tamamı zaten veritabanında ekli.")
                self.filter_data_by_license(lisans_adi)
                return

            df_new = pd.DataFrame(new_records)
            df_new.to_sql(self.table_name, conn, if_exists="append", index=False)

        info_msg = f"✓ '{lisans_adi}' lisansına {len(new_records)} yeni ders eklendi."
        if skipped_count > 0:
            info_msg += f"\n({skipped_count} ders zaten kayıtlı olduğu için atlandı)"

        self.lbl_status.config(text=info_msg.replace("\n", " "), foreground="green")
        messagebox.showinfo("İşlem Tamamlandı", info_msg)
        self.filter_data_by_license(lisans_adi)

    def delete_selected_records(self):
        """Tablodan seçilen TEK veya ÇOKLU ders kayıtlarını veritabanından topluca siler."""
        selected_items = self.tree.selection()
        
        if not selected_items:
            messagebox.showwarning("Uyarı", "Lütfen tablodan silmek istediğiniz ders kaydını (veya kayıtlarını) seçin.")
            return

        count = len(selected_items)
        confirm = messagebox.askyesno(
            "Çoklu Silme Onayı", 
            f"Seçilen {count} adet ders kaydı veritabanından kalıcı olarak silinecektir.\n\nEmin misiniz?"
        )

        if confirm:
            to_delete = []
            for item in selected_items:
                values = self.tree.item(item, "values")
                to_delete.append((values[0], values[1]))

            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.executemany(
                    f"DELETE FROM {self.table_name} WHERE lisans_adi = ? AND ders_kodu_ve_adi = ?",
                    to_delete
                )
                conn.commit()

            for item in selected_items:
                self.tree.delete(item)

            self.lbl_status.config(text=f"✓ Toplam {count} adet ders kaydı silindi.", foreground="red")
            messagebox.showinfo("Silindi", f"{count} adet ders kaydı veritabanından başarıyla silindi.")

    def export_to_excel(self):
        """Veritabanındaki tüm verileri Excel (.xlsx) dosyası olarak kaydeder."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (self.table_name,))
            if not cursor.fetchone():
                messagebox.showwarning("Uyarı", "Dışa aktarılacak veri bulunamadı.")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Dosyası", "*.xlsx"), ("Tüm Dosyalar", "*.*")],
                title="Excel Olarak Kaydet"
            )
            if not file_path:
                return

            try:
                df = pd.read_sql_query(f"SELECT * FROM {self.table_name}", conn)
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Başarılı", f"Veriler Excel dosyasına aktarıldı:\n{file_path}")
                self.lbl_status.config(text=f"✓ Excel'e aktarıldı: {file_path}", foreground="blue")
            except ModuleNotFoundError:
                messagebox.showerror(
                    "Eksik Kütüphane", 
                    "Excel çıktısı için 'openpyxl' kütüphanesi gerekli.\nTerminalde: pip install openpyxl"
                )
            except Exception as e:
                messagebox.showerror("Hata", f"Excel dosyası oluşturulurken hata oluştu:\n{e}")

    def export_to_csv(self):
        """Veritabanındaki tüm verileri CSV dosyası olarak kaydeder."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (self.table_name,))
            if not cursor.fetchone():
                messagebox.showwarning("Uyarı", "Dışa aktarılacak veri bulunamadı.")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV Dosyası", "*.csv"), ("Tüm Dosyalar", "*.*")],
                title="CSV Olarak Kaydet"
            )
            if not file_path:
                return

            try:
                df = pd.read_sql_query(f"SELECT * FROM {self.table_name}", conn)
                df.to_csv(file_path, index=False, encoding="utf-8-sig")
                messagebox.showinfo("Başarılı", f"Veriler CSV dosyasına aktarıldı:\n{file_path}")
                self.lbl_status.config(text=f"✓ CSV'ye aktarıldı: {file_path}", foreground="blue")
            except Exception as e:
                messagebox.showerror("Hata", f"CSV dosyası oluşturulurken hata oluştu:\n{e}")

    def populate_treeview(self, df: pd.DataFrame):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def load_all_data_from_db(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (self.table_name,))
            if cursor.fetchone():
                df = pd.read_sql_query(f"SELECT * FROM {self.table_name}", conn)
                self.populate_treeview(df)
                self.lbl_status.config(text=f"Toplam {len(df)} ders kaydı listeleniyor.", foreground="black")
            else:
                self.lbl_status.config(text="Veritabanında henüz kayıtlı tablo yok.", foreground="black")

    def filter_data_by_license(self, lisans_adi: str):
        with sqlite3.connect(self.db_path) as conn:
            df = pd.read_sql_query(
                f"SELECT * FROM {self.table_name} WHERE lisans_adi = ?", conn, params=(lisans_adi,)
            )
            self.populate_treeview(df)


if __name__ == "__main__":
    root = tk.Tk()
    app = MatrixGUIApp(root)
    root.mainloop()