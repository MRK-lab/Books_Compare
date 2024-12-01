import pandas as pd
import tkinter as tk
from tkinter import messagebox

# Türkçe karakterleri küçük harfe çeviren fonksiyon
def turkce_kucuk_harfe_cevir(metin):
    harf_tablosu = {
        "I": "ı", "İ": "i",
        "Ç": "ç", "Ş": "ş",
        "Ü": "ü", "Ğ": "ğ",
        "Ö": "ö", "Z": "z",
        "A": "a", "B": "b",
        "C": "c", "D": "d",
        "E": "e", "F": "f",
        "G": "g", "H": "h",
        "J": "j", "K": "k",
        "L": "l", "M": "m",
        "N": "n", "O": "o",
        "P": "p", "R": "r",
        "S": "s", "T": "t",
        "U": "u", "V": "v",
        "Y": "y", "X": "x"
    }
    return ''.join(harf_tablosu.get(harf, harf) for harf in str(metin))

# İşlemi gerçekleştiren fonksiyon
def calistir():
    try:
        # Excel dosyasının yolunu belirleyin
        excel_file = "veriler.xlsx"  # Excel dosyanızın adı

        # X ve Y verilerini ayrı sayfalardan okuyun
        x_verileri = pd.read_excel(excel_file, sheet_name="X Verileri", header=None)
        y_verileri = pd.read_excel(excel_file, sheet_name="Y Verileri", header=None)

        # Verileri küçük harfe çevirip boşlukları temizleyerek karşılaştır
        x_verileri["A Temiz"] = x_verileri[0].str.strip().apply(turkce_kucuk_harfe_cevir)
        y_verileri["A Temiz"] = y_verileri[0].str.strip().apply(turkce_kucuk_harfe_cevir)

        # X verilerinde olup Y verilerinde olmayanları bulun
        eksik_veriler = x_verileri[~x_verileri["A Temiz"].isin(y_verileri["A Temiz"])]

        # Tüm satırları yazdırmak için tüm sütunları yazıyoruz
        # Eksik veriler, A sütununda yer alan verilere karşılık gelen tüm satırları içerir.
        eksik_veriler_sonuc = eksik_veriler.drop(columns=["A Temiz"])  # "A Temiz" sütununu kaldırıyoruz

        # Sonuçları yeni bir Excel sayfasına yaz (Eksik verilerin tamamını yazıyoruz)
        with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            eksik_veriler_sonuc.to_excel(writer, sheet_name="Sonuçlar", index=False, header=False)

        # İşlem başarıyla tamamlandı mesajı
        messagebox.showinfo("Tamamlandı", "Karşılaştırma tamamlandı. Sonuçlar 'Sonuçlar' sayfasına eklendi.")
    except Exception as e:
        # Hata mesajı
        messagebox.showerror("Hata", f"Bir hata oluştu: {e}")

# Tkinter arayüzü oluşturma
root = tk.Tk()
root.title("Excel Karşılaştırma Programı")
root.geometry("400x200")

# Bilgilendirme etiketi
etiket = tk.Label(root, text="Excel dosyasını karşılaştırmak için 'Çalıştır' butonuna basın.", wraplength=350, justify="center")
etiket.pack(pady=20)

# Çalıştır butonu
buton = tk.Button(root, text="Çalıştır", command=calistir, bg="blue", fg="white", font=("Arial", 12))
buton.pack(pady=10)

# Ana döngüyü başlat
root.mainloop()
