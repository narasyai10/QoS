# 📊 Analisis Quality of Service (QoS) Jaringan menggunakan Python & TShark

Aplikasi ini dirancang untuk menganalisis **Quality of Service (QoS)** jaringan komputer menggunakan file hasil tangkapan paket (`.pcapng`) yang dihasilkan oleh Wireshark.  
Proses analisis dilakukan dengan memanfaatkan **TShark (Command Line Interface Wireshark)** dan diolah menggunakan bahasa pemrograman **Python** sehingga menghasilkan data QoS yang terstruktur dan terukur.

---

## ✨ Fitur
- Pemrosesan batch untuk beberapa file hasil capture Wireshark
- Perhitungan otomatis parameter QoS:
  - Throughput (kbps)
  - Delay (ms)
  - Jitter (ms)
  - Packet Loss (%) *(placeholder)*
- Perhitungan nilai rata-rata setiap parameter
- Ekspor hasil ke file Excel (.xlsx)
- Output langsung pada terminal untuk peninjauan cepat

---

## 🧰 Teknologi yang Digunakan
- Python 3.x
- Wireshark (TShark)
- Pandas
- Subprocess
- Microsoft Excel (output)

---

## 📁 Struktur Proyek

QoS-Network-Analysis/  
├── Simulasi01.pcapng  
├── Simulasi02.pcapng  
├── Simulasi03.pcapng  
├── Simulasi04.pcapng  
├── Simulasi05.pcapng  
├── qos_analysis.py  
└── Hasil_QoS.xlsx  

---

## ⚙️ Kebutuhan Sistem
- Python versi 3.x  
- Wireshark (termasuk TShark)  
- Library Pandas  

Instalasi dependensi:
pip install pandas

Pastikan lokasi TShark pada script sudah sesuai:
TSHARK = r"C:\Program Files\Wireshark\tshark.exe"

Untuk Linux/macOS, sesuaikan path (misalnya: /usr/bin/tshark).

---

## ▶️ Cara Menjalankan Program
1. Letakkan semua file `.pcapng` dalam satu direktori dengan file Python
2. Buka terminal atau command prompt
3. Jalankan perintah berikut:
python qos_analysis.py

4. Program akan:
- Membaca file capture paket
- Menghitung parameter QoS
- Menampilkan hasil pada terminal
- Menyimpan hasil analisis ke dalam file Excel

---

## 📊 Perhitungan Parameter QoS

### Throughput (kbps)
Throughput = (Total ukuran paket × 8) / Durasi / 1000

### Delay (ms)
Delay = Total delay / Jumlah paket yang diterima

### Jitter (ms)
Jitter = Total variasi delay / Jumlah paket yang diterima

### Packet Loss (%)
Packet loss **tidak dihitung secara langsung** dari file capture dan saat ini diset bernilai:
0.0  

Nilai ini digunakan sebagai *placeholder* untuk pengembangan selanjutnya.

---

## 📈 Output
- Output pada terminal dalam bentuk tabel
- File Excel yang dihasilkan secara otomatis:  
Hasil_QoS.xlsx

Kolom output:
- Simulasi
- Throughput (kbps)
- Delay (ms)
- Jitter (ms)
- Packet Loss (%)
- Nilai rata-rata

---

## ⚠️ Catatan
- Minimal diperlukan **2 paket** untuk menghitung delay dan jitter
- File capture dengan jumlah paket yang tidak mencukupi akan menghasilkan nilai 0
- Perhitungan packet loss memerlukan analisis berbasis sequence number atau perbandingan pengirim–penerima
- Cocok digunakan untuk:
  - Analisis performa jaringan
  - Evaluasi Quality of Service (QoS)
  - Penelitian akademik (tugas akhir / skripsi)

---

## 🎓 Konteks Akademik
Aplikasi ini dapat digunakan sebagai alat bantu dalam **penelitian jaringan komputer**, khususnya untuk menganalisis parameter Quality of Service (QoS) seperti throughput, delay, dan jitter berdasarkan data nyata hasil tangkapan paket menggunakan Wireshark.

---

## 🚀 Pengembangan Selanjutnya
- Perhitungan packet loss berbasis sequence number
- Penyaringan protokol (TCP / UDP)
- Visualisasi data dalam bentuk tabel
- Deteksi otomatis path TShark lintas platform

---

## 👤 Penulis
**Syaiful Ulum**  
Dikembangkan untuk keperluan analisis jaringan dan penelitian akademik.