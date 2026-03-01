import subprocess
import pandas as pd
import os

TSHARK = r"C:\Program Files\Wireshark\tshark.exe"

FILES = [
    "Simulasi01.pcapng",
    "Simulasi02.pcapng",
    "Simulasi03.pcapng",
    "Simulasi04.pcapng",
    "Simulasi05.pcapng"
]

def hitung_qos(file):
    cmd = [
        TSHARK,
        "-r", file,
        "-T", "fields",
        "-e", "frame.time_epoch",
        "-e", "frame.len"
    ]

    result = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )

    lines = result.stdout.strip().split("\n")

    waktu = []
    ukuran = []

    for line in lines:
        try:
            t, s = line.split("\t")
            waktu.append(float(t))
            ukuran.append(int(s))
        except:
            continue

    total_paket = len(waktu)
    print(f"{file} -> total paket terbaca: {total_paket}")

    if total_paket < 2:
        return 0, 0, 0, 0

    # =====================
    # DELAY (ms)
    # Delay = total delay / total paket diterima
    # =====================
    delay_per_paket = [(waktu[i] - waktu[i-1]) * 1000 for i in range(1, total_paket)]
    total_delay = sum(delay_per_paket)
    delay = total_delay / total_paket

    # =====================
    # JITTER (ms)
    # Jitter = total variasi delay / total paket diterima
    # =====================
    variasi_delay = [
        abs(delay_per_paket[i] - delay_per_paket[i-1])
        for i in range(1, len(delay_per_paket))
    ]
    total_variasi_delay = sum(variasi_delay)
    jitter = total_variasi_delay / total_paket

    # =====================
    # THROUGHPUT (kbps)
    # =====================
    durasi = max(waktu) - min(waktu)
    throughput = (sum(ukuran) * 8 / durasi) / 1000 if durasi > 0 else 0

    # =====================
    # PACKET LOSS (%)
    # (tidak dihitung dari Wireshark)
    # =====================
    packet_loss = 0.0

    return (
        round(throughput, 2),
        round(delay, 2),
        round(jitter, 2),
        packet_loss
    )


# =====================
# MAIN
# =====================
hasil = []

for i, f in enumerate(FILES, start=1):
    t, d, j, l = hitung_qos(f)
    hasil.append({
        "Simulasi": i,
        "Throughput (kbps)": t,
        "Delay (ms)": d,
        "Jitter (ms)": j,
        "Packet Loss (%)": l
    })

df = pd.DataFrame(hasil)

rata_rata = {
    "Simulasi": "Rata-rata",
    "Throughput (kbps)": round(df["Throughput (kbps)"].mean(), 2),
    "Delay (ms)": round(df["Delay (ms)"].mean(), 2),
    "Jitter (ms)": round(df["Jitter (ms)"].mean(), 2),
    "Packet Loss (%)": 0
}

df = pd.concat([df, pd.DataFrame([rata_rata])], ignore_index=True)

print("\n===== HASIL QUALITY OF SERVICE =====\n")
print(df)

output = os.path.join(os.getcwd(), "Hasil_QoS.xlsx")
df.to_excel(output, index=False)

print(f"\nFile Excel berhasil disimpan di:\n{output}")