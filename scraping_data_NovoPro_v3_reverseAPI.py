import requests
import pandas as pd
import time
import random
import os

# ==========================
# CONFIG
# ==========================
INPUT_FILE = "NovoPro_input_5k.xlsx"
OUTPUT_FILE = "NovoPro_output.xlsx"
AUTO_SAVE_EVERY = 20
MAX_RETRY = 3
DELAY_MIN = 1.5
DELAY_MAX = 2.5

url = "https://www.novoprolabs.com/plus/ppc.php"

headers = {
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Origin": "https://www.novoprolabs.com",
    "Referer": "https://www.novoprolabs.com/tools/convert-peptide-to-smiles-string",
    "X-Requested-With": "XMLHttpRequest",
    "User-Agent": "Mozilla/5.0"
}

# ==========================
# LOAD INPUT
# ==========================
df = pd.read_excel(INPUT_FILE)
id_smiles = df["ID"].dropna().tolist()

# ==========================
# RESUME LOGIC
# ==========================
if os.path.exists(OUTPUT_FILE):
    old_df = pd.read_excel(OUTPUT_FILE)
    done_ids = set(old_df["ID"].tolist())
    results = old_df.to_dict("records")
    print(f"🔁 Resume: {len(done_ids)} dòng đã xử lý")
else:
    done_ids = set()
    results = []
    print("🆕 Chạy mới hoàn toàn")

session = requests.Session()

print(f"Tổng số cần xử lý: {len(id_smiles)}")

for i, sid in enumerate(id_smiles, start=1):

    if sid in done_ids:
        continue

    print(f"{i}/{len(id_smiles)} → {sid}")

    payload = {
        "sr": "psmi",
        "seq": sid,
        "ct": "linear",
        "p": ""
    }

    success = False

    for attempt in range(1, MAX_RETRY + 1):
        try:
            response = session.post(
                url,
                headers=headers,
                data=payload,
                timeout=30
            )
            if response.status_code == 200:
                data = response.json()
                if isinstance(data, list) and len(data) > 1:
                    result_text = data[1][0][1]
                else:
                    result_text = "INVALID_RESPONSE"
                success = True
                break
            else:
                result_text = f"HTTP_{response.status_code}"
        except Exception as e:
            print(f"⚠ Retry {attempt}/{MAX_RETRY}")
            time.sleep(2)
    if not success:
        result_text = "ERROR"

    results.append({
        "ID": sid,
        "Code SMILES": result_text
    })

    # Auto save thường xuyên để tránh mất dữ liệu
    if i % AUTO_SAVE_EVERY == 0:
        pd.DataFrame(results).to_excel(OUTPUT_FILE, index=False)
        print("💾 Auto saved")

    # Delay chống block
    time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

# Save cuối cùng
pd.DataFrame(results).to_excel(OUTPUT_FILE, index=False)

print("🎉 HOÀN TẤT!")