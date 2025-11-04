# -*- coding: utf-8 -*-
"""
Created on Tue Nov  4 14:23:03 2025

@author: miqdad
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scrape review Google Play untuk banyak APK sekaligus yang ada di file Excel.

- Input  : Updated_List_APK.xlsx
  - Kolom wajib:
        "Alamat APK  Android (com.xxx.xx)"  -> app_id / package
  - Kolom opsional (tapi disarankan):
        "Nama Platform"                     -> nama platform / apk_name

- Output : reviews_all_apk.xlsx
  - Beberapa sheet:
      - Jika total baris <= 1.000.000  -> 1 sheet  : "reviews"
      - Jika > 1.000.000              -> multi-sheet: "reviews_1", "reviews_2", ...
  - Semua APK digabung
  - Kolom tambahan "apk_name"
  - Tanggal (at, replyAt) hanya YYYY-MM-DD
  - Jika satu APK gagal / tidak ada review → dilewati, lanjut ke APK berikutnya
"""

import sys
import time
import re
import math
from datetime import datetime

import pandas as pd
from google_play_scraper import reviews, Sort

# ====================== KONFIGURASI ======================

# Nama file Excel yang berisi daftar APK
APK_LIST_FILE = "Updated_List_APK.xlsx"

# Nama file output review
OUTPUT_FILE = "reviews_all_apk.xlsx"

# Bahasa dan negara Play Store
LANG = "id"          # contoh: "id", "en"
COUNTRY = "id"       # contoh: "id", "us"

# Batasan scraping
BATCH_SIZE = 100          # jumlah review per request (maks 200)
MAX_REVIEWS_PER_APP = None   # None = ambil semua review yang tersedia untuk tiap APK

# Jeda antar request (supaya tidak terlalu agresif)
RATE_LIMIT_SEC = 1.0

# Nama kolom di Excel input
COL_PLATFORM_NAME = "Nama Platform"
COL_APP_ID = "Alamat APK  Android (com.xxx.xx)"

# Batas aman baris per sheet (Excel .xlsx max 1.048.576)
MAX_ROWS_PER_SHEET = 1_000_000

# =========================================================

DANGEROUS_PREFIX = ('=', '+', '-', '@')


def normalize_text(val):
    """
    Bersihkan newline/tab berlebih dan lindungi dari formula injection Excel.
    """
    if val is None:
        return None
    s = str(val)
    # hapus newline/tab berlebih
    s = re.sub(r'\s+', ' ', s).strip()
    # lindungi jika diawali dengan karakter formula Excel
    if s.startswith(DANGEROUS_PREFIX):
        s = "'" + s
    return s


def to_date_str(dt):
    """
    Ubah datetime ke string tanggal YYYY-MM-DD (tanpa jam).
    Kalau bukan datetime, balikin str(dt).
    """
    if dt is None:
        return None
    try:
        return dt.date().isoformat()
    except Exception:
        return str(dt)


def load_apk_list_from_excel(path):
    """
    Baca daftar APK dari file Excel.
    Mengembalikan list of (app_id, apk_name).
    """
    try:
        df_list = pd.read_excel(path)
    except FileNotFoundError:
        print(f"[FATAL] File Excel '{path}' tidak ditemukan. "
              f"Letakkan {path} di folder yang sama dengan script ini.")
        sys.exit(1)

    if COL_APP_ID not in df_list.columns:
        raise ValueError(
            f"Kolom '{COL_APP_ID}' tidak ditemukan di {path}. "
            f"Kolom yang tersedia: {list(df_list.columns)}"
        )

    has_name = COL_PLATFORM_NAME in df_list.columns

    apk_pairs = []
    for _, row in df_list.iterrows():
        app_id = row[COL_APP_ID]
        if pd.isna(app_id):
            continue
        app_id = str(app_id).strip()
        if not app_id:
            continue

        if has_name:
            apk_name = str(row[COL_PLATFORM_NAME]).strip() or app_id
        else:
            apk_name = app_id

        apk_pairs.append((app_id, apk_name))

    if not apk_pairs:
        raise ValueError(f"Tidak ada app_id valid di file {path}")

    print(f"Ditemukan {len(apk_pairs)} APK di '{path}'")
    return apk_pairs


def fetch_reviews_for_app(app_id, apk_name,
                          lang=LANG,
                          country=COUNTRY,
                          batch=BATCH_SIZE,
                          max_reviews=MAX_REVIEWS_PER_APP,
                          rate_limit=RATE_LIMIT_SEC):
    """
    Ambil review untuk satu APK.
    Kalau error / tidak ada review → return [].
    """
    print(f"\n=== Mulai scrape: {apk_name} ({app_id}) ===")
    sys.stdout.flush()

    all_rows = []
    continuation_token = None
    fetched = 0

    while True:
        # hitung berapa review yang mau diambil di batch ini
        count = batch
        if max_reviews is not None:
            remain = max_reviews - fetched
            if remain <= 0:
                break
            count = min(count, remain)

        try:
            result, continuation_token = reviews(
                app_id,
                lang=lang,
                country=country,
                sort=Sort.NEWEST,
                count=count,
                continuation_token=continuation_token
            )
        except Exception as e:
            print(f"[ERROR] Gagal mengambil review untuk {apk_name} ({app_id}): {e}")
            print("-> Lewati APK ini dan lanjut ke berikutnya.")
            return []

        if not result:
            if fetched == 0:
                print(f"Tidak ada review untuk {apk_name} ({app_id}).")
            break

        for r in result:
            all_rows.append({
                "apk_name": apk_name,
                "reviewId": r.get("reviewId"),
                "userName": r.get("userName"),
                "score": r.get("score"),
                "text": r.get("content") or r.get("summary") or "",
                "at": to_date_str(r.get("at")),
                "replyText": r.get("replyContent") or "",
                "replyAt": to_date_str(r.get("repliedAt")),
                "thumbsUpCount": r.get("thumbsUpCount", 0),
                "version": r.get("reviewCreatedVersion") or r.get("version"),
            })
            fetched += 1

        print(f"[{datetime.now().isoformat()}] {apk_name} - total fetched: {fetched} (+{len(result)})")
        sys.stdout.flush()

        if not continuation_token:
            break
        if max_reviews is not None and fetched >= max_reviews:
            break

        time.sleep(rate_limit)

    return all_rows


def save_to_xlsx(df, path):
    """
    Simpan DataFrame ke Excel dengan multi-sheet jika baris > MAX_ROWS_PER_SHEET.
    """
    # urutan kolom yang diinginkan
    cols = [
        "apk_name",
        "reviewId",
        "userName",
        "score",
        "text",
        "at",
        "replyText",
        "replyAt",
        "thumbsUpCount",
        "version",
    ]
    df = df[[c for c in cols if c in df.columns]].copy()

    # normalisasi teks untuk kolom string
    for c in df.columns:
        if pd.api.types.is_object_dtype(df[c].dtype):
            df[c] = df[c].map(normalize_text)

    # sort berdasarkan tanggal (kalau ada)
    if "at" in df.columns:
        df = df.sort_values("at", ascending=False)

    n_rows = len(df)
    n_sheets = max(1, math.ceil(n_rows / MAX_ROWS_PER_SHEET))

    print(f"\nMenyimpan {n_rows:,} baris ke '{path}' dalam {n_sheets} sheet ...")

    # set lebar kolom
    widths = {
        0: 25,  # apk_name
        1: 22,  # reviewId
        2: 22,  # userName
        3: 8,   # score
        4: 80,  # text
        5: 14,  # at
        6: 80,  # replyText
        7: 14,  # replyAt
        8: 14,  # thumbsUpCount
        9: 14,  # version
    }

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        for i in range(n_sheets):
            start = i * MAX_ROWS_PER_SHEET
            end = min((i + 1) * MAX_ROWS_PER_SHEET, n_rows)
            chunk = df.iloc[start:end]

            if n_sheets == 1:
                sheet_name = "reviews"
            else:
                sheet_name = f"reviews_{i + 1}"

            print(f"  - Sheet {sheet_name}: baris {start + 1} s.d. {end}")
            chunk.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.sheets[sheet_name]
            for idx, col in enumerate(chunk.columns):
                ws.set_column(idx, idx, widths.get(idx, 18))

    print(f"\nSelesai. Tersimpan {n_rows:,} baris ke file '{path}'.")


def main():
    print("=== CONFIG ===")
    print("Excel APK list :", APK_LIST_FILE)
    print("Output file    :", OUTPUT_FILE)
    print("Lang           :", LANG)
    print("Country        :", COUNTRY)
    print("Batch size     :", BATCH_SIZE)
    print("Max per app    :", MAX_REVIEWS_PER_APP)
    print("Rate (sec)     :", RATE_LIMIT_SEC)
    print("Max rows/sheet :", MAX_ROWS_PER_SHEET)
    print("================\n")

    # 1) Baca daftar APK dari Excel
    apk_pairs = load_apk_list_from_excel(APK_LIST_FILE)

    # 2) Loop setiap APK dan kumpulkan semua review
    all_rows = []
    for app_id, apk_name in apk_pairs:
        rows = fetch_reviews_for_app(
            app_id=app_id,
            apk_name=apk_name,
            lang=LANG,
            country=COUNTRY,
            batch=BATCH_SIZE,
            max_reviews=MAX_REVIEWS_PER_APP,
            rate_limit=RATE_LIMIT_SEC
        )

        if not rows:
            print(f"-> Tidak ada review / gagal untuk {apk_name} ({app_id}). Lanjut ke APK berikutnya.\n")
            continue

        all_rows.extend(rows)

    if not all_rows:
        print("\nTidak ada review yang berhasil dikumpulkan dari semua APK.")
        return

    # 3) Gabungkan ke DataFrame
    df = pd.DataFrame(all_rows)

    # Deduplikasi (apk_name, reviewId) kalau memungkinkan
    if "reviewId" in df.columns:
        subset = ["reviewId"]
        if "apk_name" in df.columns:
            subset = ["apk_name", "reviewId"]
        before = len(df)
        df = df.drop_duplicates(subset=subset)
        after = len(df)
        if after < before:
            print(f"Dedupe: {before} -> {after} baris (berdasarkan {subset})")

    # 4) Simpan ke Excel (multi-sheet jika perlu)
    save_to_xlsx(df, OUTPUT_FILE)


if __name__ == "__main__":
    main()
