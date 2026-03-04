import struct
from pathlib import Path

import pandas as pd

import tkinter as tk
from tkinter import filedialog, messagebox

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
    0x36: "PD",
    0x37: "PD_DZ"
}

# Возможные имена колонок (на случай разных логгеров/версий)
DATA_COL_CANDIDATES = ["Data", "DATA", "Data_Hex", "DATA_HEX", "DataHex", "Bytes", "Payload"]
TS_COL_CANDIDATES = ["Timestamp", "TIME", "Time", "DateTime", "Datetime", "TimeStamp"]


def bytes_list_to_float_le(hex_bytes_4) -> float:
    b = bytes(int(x, 16) for x in hex_bytes_4)
    return struct.unpack("<f", b)[0]


def find_column(df: pd.DataFrame, candidates) -> str | None:
    # нормализуем имена колонок (убираем пробелы, приводим к нижнему)
    norm = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in norm:
            return norm[key]
    return None


def merge_csv_files(files: list[str]) -> pd.DataFrame:
    frames = []
    for f in files:
        # utf-8-sig убирает BOM (﻿ в начале)
        df = pd.read_csv(f, encoding="utf-8-sig")
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def convert_dataframe_to_excel(df: pd.DataFrame, output_path: Path):
    data_col = find_column(df, DATA_COL_CANDIDATES)
    ts_col = find_column(df, TS_COL_CANDIDATES)

    if not data_col or not ts_col:
        raise ValueError(
            "Не нашёл нужные колонки.\n"
            "Ожидаю Data/Data_Hex и Timestamp.\n\n"
            f"Найденные колонки: {list(df.columns)}"
        )

    ts = pd.to_datetime(df[ts_col], errors="coerce")

    rows = []
    for data_str, t in zip(df[data_col].astype(str), ts):
        if pd.isna(t):
            continue

        parts = data_str.strip().split()
        if len(parts) < 5:
            continue

        msg_id = int(parts[0], 16)
        if msg_id not in PID_IDS:
            continue

        value = bytes_list_to_float_le(parts[-4:])
        rows.append({"Timestamp": t, "Signal": PID_IDS[msg_id], "Value": value})

    decoded = pd.DataFrame(rows).sort_values("Timestamp")
    if decoded.empty:
        raise ValueError("Нет декодируемых PID-сообщений (0x27..0x37) в выбранных файлах.")

    # время в секундах от старта
    t0 = decoded["Timestamp"].iloc[0]
    decoded["TimeSec"] = (decoded["Timestamp"] - t0).dt.total_seconds()

    # таблица по времени + заполнение последним известным значением
    wide = decoded.pivot_table(
        index="TimeSec", columns="Signal", values="Value", aggfunc="last"
    ).sort_index().ffill()

    out_df = wide.reset_index()

    # пишем Excel (без графиков)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="data", index=False)


def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)  # чтобы окна не прятались за cmd/VSCode

    try:
        files = filedialog.askopenfilenames(
            parent=root,
            title="Выберите CSV файлы логгера (можно несколько)",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not files:
            return

        # Автосортировка по имени (095133_000, 095133_001, ...)
        files = sorted(files)

        save_path = filedialog.asksaveasfilename(
            parent=root,
            title="Сохранить итоговый Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="pid_merged.xlsx"
        )
        if not save_path:
            return

        merged = merge_csv_files(list(files))
        convert_dataframe_to_excel(merged, Path(save_path))

        messagebox.showinfo("Готово", f"Файл создан:\n{save_path}", parent=root)

    except Exception as e:
        messagebox.showerror("Ошибка", str(e), parent=root)

    finally:
        # ВАЖНО: иначе процесс может не завершиться
        root.destroy()


if __name__ == "__main__":
    main()

