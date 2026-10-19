"""Aplikasi GUI Rekonsiliasi Pajak BKU SKPD SIPD + Coretax."""
import os
import re
import subprocess
import sys
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import konsolidasi as ks

if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # teks tajam di layar ber-skala
    except Exception:
        pass

NAMA_APP = "Rekonsiliasi Pajak BKU SKPD SIPD + Coretax"
VERSI = "1.1"
TANGGAL_RILIS = "Oktober 2026"
PEMBUAT = "JK"

TIPE_BKU = [("Excel", "*.xlsx *.xlsm"), ("CSV", "*.csv"), ("Semua file", "*.*")]
TIPE_CORETAX = [("CSV", "*.csv"), ("Excel", "*.xlsx"), ("Semua file", "*.*")]

# (teks, gaya) untuk jendela Petunjuk Penggunaan
PETUNJUK = [
    ("Petunjuk Penggunaan", "judul"),
    ("Aplikasi ini mencocokkan potongan pajak di BKU SKPD (unduhan SIPD) dengan setoran di Buku Besar Coretax: "
     "NTPN BKU dicocokkan dengan ReferenceNumber Coretax, dan kolom Penerimaan BKU dengan kolom Amount Coretax.", "isi"),
    ("Langkah-langkah", "bagian"),
    ("1. Unduh BKU pajak SKPD dari SIPD dalam format Excel (.xlsx).", "butir"),
    ("2. Unduh Buku Besar dari Coretax dalam format CSV. Jangan buka lalu simpan ulang CSV ini di Excel, "
     "karena angka panjang (NPWP, NTPN) bisa rusak.", "butir"),
    ("3. Di bagian '1. Upload file', klik 'Pilih File...' untuk BKU SKPD dan untuk Buku Besar Coretax.", "butir"),
    ("4. Klik 'Proses Rekonsiliasi'. Ringkasan dan daftar perbaikan tampil di bagian '2. Hasil rekonsiliasi'.", "butir"),
    ("5. Jendela simpan muncul otomatis. Klik Save; file Excel hasil langsung terbuka.", "butir"),
    ("6. Tombol 'Simpan Excel...', 'Buka File Hasil' dan 'Buka Folder' bisa dipakai setelah proses selesai.", "butir"),
    ("Kolom yang dibutuhkan", "bagian"),
    ("BKU SKPD: NTPN, Uraian, Penerimaan, Pengeluaran (wajib). No, Tanggal, Nomor Bukti, Nama Rekening, "
     "ID Billing dan Kode Rekening dipakai bila ada.", "butir"),
    ("Buku Besar Coretax: ReferenceNumber dan Amount (wajib). AmountLeft, RevenueCode, PaymentCode, PeriodCode, "
     "TransactionDate, PostingDate, IsCancelled dan ReferenceType dipakai bila ada.", "butir"),
    ("Arti status", "bagian"),
    ("COCOK (hijau): NTPN, nilai dan tanggal sama.", "butir"),
    ("COCOK - TANGGAL BEDA (kuning): NTPN dan nilai sama, tanggal BKU beda dengan tanggal setor.", "butir"),
    ("NIHIL (abu-abu): potongan Rp0 tanpa NTPN, tidak perlu tindakan.", "butir"),
    ("NTPN SALAH KETIK (merah): NTPN BKU beda maksimal 3 karakter dengan NTPN Coretax; NTPN yang benar ditampilkan.",
     "butir"),
    ("NILAI BEDA (merah): NTPN sama, nilai Penerimaan BKU berbeda dengan Amount Coretax.", "butir"),
    ("TIDAK DITEMUKAN (merah): NTPN BKU tidak ada di Coretax, atau kosong padahal ada potongan.", "butir"),
    ("NTPN GANDA DI BKU (merah): NTPN yang sama dipakai lebih dari satu transaksi BKU.", "butir"),
    ("NTPN DIBATALKAN DI CORETAX (merah): NTPN berstatus dibatalkan di Coretax.", "butir"),
    ("Isi file Excel hasil", "bagian"),
    ("Kesimpulan: ringkasan angka, hasil pencocokan, daftar yang perlu diperbaiki (dengan nomor bukti 6 digit), catatan.",
     "butir"),
    ("Rekonsiliasi NTPN: semua transaksi BKU berdampingan dengan data Coretax, diberi warna sesuai status.", "butir"),
    ("Perlu Perbaikan: hanya transaksi yang perlu dikoreksi beserta tindakannya.", "butir"),
    ("Coretax Tidak di BKU: NTPN yang ada di Coretax tetapi tidak tercatat di BKU.", "butir"),
    ("Konsolidasi Jenis Pajak: total per jenis pajak dan ringkasan per KAP-KJS.", "butir"),
    ("Masalah yang sering muncul", "bagian"),
    ("'Format file tidak sesuai': periksa apakah file BKU dan Coretax tertukar, atau kolom wajib tidak ada.", "butir"),
    ("'Tidak bisa menyimpan': file hasil dengan nama yang sama sedang terbuka di Excel. Tutup dulu, lalu simpan lagi.",
     "butir"),
    ("File .xls tidak bisa dibaca: buka di Excel lalu Save As .xlsx.", "butir"),
    ("NPWP di judul hasil tampil '-': CSV Coretax pernah disimpan ulang di Excel. Unduh ulang CSV asli.", "butir"),
]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{NAMA_APP} - Versi {VERSI}")
        self.geometry("960x680")
        self.minsize(760, 540)
        self.var_bku = tk.StringVar()
        self.var_ct = tk.StringVar()
        self.var_status = tk.StringVar(value="Pilih file BKU SKPD dan Buku Besar Coretax, lalu klik Proses Rekonsiliasi.")
        self.hasil = None
        self.file_hasil = None
        self._bangun()

    # ------------------------------------------------------------------ tampilan
    def _bangun(self):
        style = ttk.Style(self)
        style.configure("Judul.TLabel", font=("Segoe UI", 16, "bold"), foreground="#1F3A6B")
        style.configure("Proses.TButton", font=("Segoe UI", 10, "bold"), padding=(18, 6))

        menubar = tk.Menu(self)
        bantuan = tk.Menu(menubar, tearoff=0)
        bantuan.add_command(label="Petunjuk Penggunaan", accelerator="F1", command=self.petunjuk)
        bantuan.add_separator()
        bantuan.add_command(label="Tentang Aplikasi", command=self.tentang)
        menubar.add_cascade(label="Bantuan", menu=bantuan)
        self.configure(menu=menubar)
        self.bind("<F1>", lambda e: self.petunjuk())

        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(5, weight=1)

        ttk.Label(root, text=NAMA_APP, style="Judul.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(root, text="Mencocokkan NTPN dan nilai Penerimaan di BKU SKPD dengan Amount di Buku Besar Coretax"
                  ).grid(row=1, column=0, sticky="w", pady=(0, 12))

        box = ttk.LabelFrame(root, text=" 1. Upload file ", padding=12)
        box.grid(row=2, column=0, sticky="ew")
        box.columnconfigure(1, weight=1)
        self._baris_file(box, 0, "BKU SKPD (.xlsx)", self.var_bku, TIPE_BKU)
        self._baris_file(box, 1, "Buku Besar Coretax (.csv)", self.var_ct, TIPE_CORETAX)

        aksi = ttk.Frame(root)
        aksi.grid(row=3, column=0, sticky="w", pady=(12, 4))
        self.btn_proses = ttk.Button(aksi, text="Proses Rekonsiliasi", style="Proses.TButton", command=self.proses)
        self.btn_proses.pack(side="left")
        self.btn_simpan = ttk.Button(aksi, text="Simpan Excel...", command=self.simpan, state="disabled")
        self.btn_simpan.pack(side="left", padx=(10, 0))
        self.btn_buka = ttk.Button(aksi, text="Buka File Hasil", command=self.buka, state="disabled")
        self.btn_buka.pack(side="left", padx=(6, 0))
        self.btn_folder = ttk.Button(aksi, text="Buka Folder", command=self.buka_folder, state="disabled")
        self.btn_folder.pack(side="left", padx=(6, 0))

        ttk.Label(root, textvariable=self.var_status, foreground="#555555").grid(row=4, column=0, sticky="w", pady=(0, 8))

        box2 = ttk.LabelFrame(root, text=" 2. Hasil rekonsiliasi ", padding=8)
        box2.grid(row=5, column=0, sticky="nsew")
        box2.rowconfigure(0, weight=1)
        box2.columnconfigure(0, weight=1)
        self.txt = tk.Text(box2, wrap="word", font=("Segoe UI", 10), relief="flat", padx=10, pady=8,
                           background="#FFFFFF", state="disabled")
        sb = ttk.Scrollbar(box2, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.txt.tag_configure("judul", font=("Segoe UI", 12, "bold"), foreground="#1F3A6B")
        self.txt.tag_configure("sub", foreground="#666666")
        self.txt.tag_configure("bagian", font=("Segoe UI", 10, "bold"), foreground="#1F3A6B", spacing1=6)
        self.txt.tag_configure("butir", lmargin1=28, lmargin2=40)
        self.txt.tag_configure("isi", lmargin1=8, lmargin2=24)
        self.txt.tag_configure("nilai", font=("Consolas", 10, "bold"))

    def _baris_file(self, parent, baris, label, var, tipe):
        ttk.Label(parent, text=label, width=26).grid(row=baris, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=baris, column=1, sticky="ew", padx=8)
        ttk.Button(parent, text="Pilih File...", command=lambda: self._pilih(var, label, tipe)).grid(row=baris, column=2)

    def _pilih(self, var, label, tipe):
        awal = next((str(Path(p).parent) for p in (var.get(), self.var_bku.get(), self.var_ct.get()) if p), str(Path.home()))
        path = filedialog.askopenfilename(title=f"Pilih file {label}", initialdir=awal, filetypes=tipe)
        if path:
            var.set(os.path.normpath(path))

    def _sibuk(self, sibuk, pesan):
        self.var_status.set(pesan)
        self.configure(cursor="watch" if sibuk else "")
        self.btn_proses.configure(state="disabled" if sibuk else "normal")
        self.update_idletasks()

    # ------------------------------------------------------------------ aksi
    def proses(self):
        bku, ct = self.var_bku.get().strip(), self.var_ct.get().strip()
        for path, nama in ((bku, "BKU SKPD"), (ct, "Buku Besar Coretax")):
            if not path or not Path(path).is_file():
                messagebox.showwarning("File belum dipilih", f"Pilih file {nama} terlebih dahulu.")
                return
        self._sibuk(True, "Memproses rekonsiliasi...")
        try:
            hasil = ks.konsolidasikan(bku, ct)
        except ValueError as e:
            self._sibuk(False, "Gagal memproses - format file tidak sesuai.")
            messagebox.showerror("Format file tidak sesuai", str(e))
            return
        except Exception:
            self._sibuk(False, "Gagal memproses.")
            messagebox.showerror("Terjadi kesalahan", traceback.format_exc())
            return

        self.hasil, self.file_hasil = hasil, None
        self._tampilkan(hasil)
        self._sibuk(False, f"Selesai: {hasil.jumlah_terima} transaksi BKU, {len(hasil.coretax_valid)} NTPN Coretax, "
                           f"{hasil.jumlah_temuan} temuan perlu diperbaiki.")
        self.btn_simpan.configure(state="normal")
        self.btn_buka.configure(state="disabled")
        self.btn_folder.configure(state="disabled")
        self.simpan()

    def simpan(self):
        if not self.hasil:
            return
        awal = Path(self.var_bku.get()).parent
        nama = "Rekonsiliasi_Pajak_" + re.sub(r"[^\w-]+", "_", self.hasil.masa).strip("_") + ".xlsx"
        path = filedialog.asksaveasfilename(title="Simpan hasil rekonsiliasi", initialdir=str(awal), initialfile=nama,
                                            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            self.var_status.set("Hasil belum disimpan. Klik 'Simpan Excel...' untuk menyimpan.")
            return
        try:
            ks.simpan_excel(self.hasil, path)
        except PermissionError:
            messagebox.showerror("Tidak bisa menyimpan",
                                 f"File {Path(path).name} sedang dibuka di Excel atau tidak bisa ditulis.\n"
                                 "Tutup file tersebut lalu klik 'Simpan Excel...' lagi.")
            return
        except Exception:
            messagebox.showerror("Gagal menyimpan", traceback.format_exc())
            return
        self.file_hasil = os.path.normpath(path)
        self.var_status.set(f"Tersimpan: {self.file_hasil}")
        self.btn_buka.configure(state="normal")
        self.btn_folder.configure(state="normal")
        self.buka()

    def buka(self):
        if self.file_hasil:
            _buka_path(self.file_hasil)

    def buka_folder(self):
        if self.file_hasil:
            _buka_path(str(Path(self.file_hasil).parent))

    # ------------------------------------------------------------------ menu bantuan
    def petunjuk(self):
        if getattr(self, "win_petunjuk", None) and self.win_petunjuk.winfo_exists():
            self.win_petunjuk.lift()
            return
        win = self.win_petunjuk = tk.Toplevel(self)
        win.title(f"Petunjuk Penggunaan - {NAMA_APP}")
        win.geometry("760x600")
        win.transient(self)
        frame = ttk.Frame(win, padding=8)
        frame.pack(fill="both", expand=True)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        txt = tk.Text(frame, wrap="word", font=("Segoe UI", 10), relief="flat", padx=12, pady=10, background="#FFFFFF")
        sb = ttk.Scrollbar(frame, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        txt.tag_configure("judul", font=("Segoe UI", 13, "bold"), foreground="#1F3A6B", spacing3=6)
        txt.tag_configure("bagian", font=("Segoe UI", 10, "bold"), foreground="#1F3A6B", spacing1=10, spacing3=2)
        txt.tag_configure("isi", spacing3=2)
        txt.tag_configure("butir", lmargin1=16, lmargin2=30, spacing3=2)
        for teks, gaya in PETUNJUK:
            txt.insert("end", teks + "\n", gaya)
        txt.configure(state="disabled")
        ttk.Button(frame, text="Tutup", command=win.destroy).grid(row=1, column=0, columnspan=2, sticky="e", pady=(8, 0))
        win.bind("<Escape>", lambda e: win.destroy())

    def tentang(self):
        messagebox.showinfo(
            "Tentang Aplikasi",
            f"{NAMA_APP}\n"
            f"Versi {VERSI} ({TANGGAL_RILIS})\n"
            f"Dibuat oleh: {PEMBUAT}\n\n"
            "Mencocokkan NTPN dan nilai potongan pajak di BKU SKPD (SIPD) dengan setoran di Buku Besar Coretax, "
            "lalu menyusun file Excel berisi kesimpulan dan daftar yang perlu diperbaiki.\n\n"
            f"Python {sys.version.split()[0]}",
            parent=self)

    def _tampilkan(self, hasil):
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        for teks, nilai, gaya in ks.susun_kesimpulan(hasil):
            self.txt.insert("end", teks, gaya)
            if nilai is not None:
                self.txt.insert("end", f"   {ks.rupiah(nilai)}", (gaya, "nilai"))
            self.txt.insert("end", "\n", gaya)
        self.txt.configure(state="disabled")
        self.txt.yview_moveto(0)


def _buka_path(path):
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


if __name__ == "__main__":
    App().mainloop()
