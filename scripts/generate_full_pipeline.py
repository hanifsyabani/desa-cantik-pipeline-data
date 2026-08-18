"""CLI entrypoint for the Desa Cantik full Excel pipeline."""

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input")
    parser.add_argument("output")
    parser.add_argument(
        "--min-harga",
        type=int,
        default=1000,
        help="Harga di bawah ini dibuang dari sheet Cleaned (default 1000)",
    )
    args = parser.parse_args()

    from desa_cantik.pipeline import generate_file

    stats = generate_file(args.input, args.output, args.min_harga)

    print(
        "Data responden terdeteksi: "
        f"{stats['n_rows']} baris "
        f"(sheet1 baris {stats['src_data_start_row']}-{stats['last_row']})"
    )
    print(f"'Dashboard (Raw)': {stats['n_raw']} baris")
    print(f"'Dashboard (Cleaned)': {stats['n_clean']} baris")
    print(f"  - dibuang (tanpa jawaban): {stats['n_blank']}")
    print(f"  - dibuang (wilayah gagal lookup / #N/A): {stats['n_dropped_wilayah']}")
    print(
        "  - dibuang "
        f"(harga < {args.min_harga}, dianggap salah ketik): {stats['n_dropped_price']}"
    )
    print(f"Selesai. File disimpan: {args.output}")


if __name__ == "__main__":
    main()
