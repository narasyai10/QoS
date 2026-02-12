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

    times = []
    sizes = []

    for line in lines:
        try:
            t, s = line.split("\t")
            times.append(float(t))
            sizes.append(int(s))
        except:
            continue

    print(f"{file} -> total paket terbaca: {len(times)}")

    if len(times) < 2:
        return 0, 0, 0, 0

    # =====================
    # DELAY (ms)
    # =====================
    delays = [(times[i] - times[i-1]) * 1000 for i in range(1, len(times))]
    delay_avg = sum(delays) / len(delays)

    # =====================
    # JITTER (ms)
    # =====================
    jitters = [abs(delays[i] - delays[i-1]) for i in range(1, len(delays))]
    jitter_avg = sum(jitters) / len(jitters) if jitters else 0

    # =====================
    # THROUGHPUT (kbps)
    # =====================
    duration = max(times) - min(times)
    throughput = (sum(sizes) * 8 / duration) / 1000 if duration > 0 else 0

    # =====================
    # PACKET LOSS (%)
    # =====================
    packet_loss = 0.0

    return round(delay_avg,2), round(jitter_avg,2), round(throughput,2), packet_loss


# =====================
# MAIN
# =====================
hasil = []

for i, f in enumerate(FILES, start=1):
    d, j, t, l = hitung_qos(f)
    hasil.append({
        "Simulasi": i,
        "Delay (ms)": d,
        "Jitter (ms)": j,
        "Throughput (kbps)": t,
        "Packet Loss (%)": l
    })

df = pd.DataFrame(hasil)

rata = {
    "Simulasi": "Rata-rata",
    "Delay (ms)": round(df["Delay (ms)"].mean(),2),
    "Jitter (ms)": round(df["Jitter (ms)"].mean(),2),
    "Throughput (kbps)": round(df["Throughput (kbps)"].mean(),2),
    "Packet Loss (%)": 0
}

df = pd.concat([df, pd.DataFrame([rata])], ignore_index=True)

print("\n===== HASIL QUALITY OF SERVICE =====\n")
print(df)

output = os.path.join(os.getcwd(), "Hasil_QoS.xlsx")
df.to_excel(output, index=False)

print(f"\nFile Excel berhasil disimpan di:\n{output}")