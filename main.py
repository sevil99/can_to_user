import pandas as pd
import struct
from pathlib import Path

INPUT_FILE = "can_messages.csv"   # или .xlsx
OUTPUT_FILE = "pid_decoded.xlsx"

PID_IDS = {
    0x27: "PV",
    0x28: "SP",
    0x29: "CV",
    0x30: "CV_P",
    0x31: "CV_I",
    0x32: "CV_D",
    0x33: "Kp",
    0x34: "Ki",
    0x35: "Kd",
}

def load_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    if path.suffix.lower() in (".xls", ".xlsx"):
        return pd.read_excel(path)
    raise ValueError("Поддерживаются только CSV и Excel")

def bytes_list_to_float_le(hex_bytes_4):
    b = bytes(int(x, 16) for x in hex_bytes_4)
    return struct.unpack("<f", b)[0]

def main():
    df = load_table(Path(INPUT_FILE))

    if "Data" not in df.columns:
        raise ValueError("Не найдена колонка 'Data'. Проверь формат входного файла.")

    result = {name: [] for name in PID_IDS.values()}

    for s in df["Data"].astype(str):
        parts = s.strip().split()
        if len(parts) < 5:
            continue

        msg_id = int(parts[0], 16)
        if msg_id not in PID_IDS:
            continue

        value = bytes_list_to_float_le(parts[-4:])
        result[PID_IDS[msg_id]].append(value)

    out_df = pd.DataFrame({k: pd.Series(v) for k, v in result.items()})
    out_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Готово: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()

