"""Inspect saved source files; never modify raw contents.

Dependencies: pandas, openpyxl, pypdf.
"""
import argparse
import importlib.metadata
import json
from pathlib import Path

import openpyxl
import pandas as pd
from pypdf import PdfReader


def csv_read(path, **kwargs):
    for encoding in ("utf-8-sig", "cp949"):
        try:
            return pd.read_csv(path, encoding=encoding, **kwargs), encoding
        except UnicodeDecodeError:
            pass
    raise ValueError(f"Unsupported source encoding: {path}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args()
    root = args.root.resolve()
    results = {"ridership": {}, "crowding": [], "coordinates": {}, "od": {}}
    dictionary = []

    path = next(root.glob("raw/01_승하차/*.csv"))
    df, encoding = csv_read(path, dtype={"역번호": str})
    keys = ["수송일자", "호선명", "역번호", "승하차구분", "승객유형"]
    values = df.iloc[:, 7:].apply(pd.to_numeric, errors="coerce")
    station_keys = df[["호선명", "역번호"]].drop_duplicates()
    dates = pd.to_datetime(df["수송일자"], errors="raise")
    r = results["ridership"]
    r.update({"file": str(path.relative_to(root)), "encoding": encoding, "rows": len(df), "columns": len(df.columns),
              "column_names": df.columns.tolist(), "date_min": str(dates.min().date()), "date_max": str(dates.max().date()),
              "date_count": dates.nunique(), "line_station_count": len(station_keys),
              "station_name_count": df["역명"].nunique(), "lines": sorted(df["호선명"].unique().tolist()),
              "passenger_types": sorted(df["승객유형"].unique().tolist()), "directions": sorted(df["승하차구분"].unique().tolist()),
              "key_duplicate_rows": int(df.duplicated(keys).sum()),
              "missing_count_cells": int(values.isna().sum().sum()), "negative_count_cells": int((values < 0).sum().sum()),
              "missing_key_cells": int(df[keys].isna().sum().sum()),
              "rows_by_month": dates.dt.strftime("%Y-%m").value_counts().sort_index().to_dict(),
              "stations_by_line": station_keys.groupby("호선명").size().to_dict(),
              "time_columns": df.columns[7:].tolist()})
    dictionary.append({"dataset": "ridership", "file": str(path.relative_to(root)), "columns": df.columns.tolist()})
    del values

    coord_path = next(root.glob("raw/03_역좌표/*.csv"))
    coords, enc = csv_read(coord_path, dtype={"고유역번호(외부역코드)": str})
    coord_keys = {(f"{line}호선", code) for line, code in zip(coords["호선"], coords["고유역번호(외부역코드)"])}
    missing_coords = station_keys[[tuple(x) not in coord_keys for x in station_keys.to_numpy()]]
    results["coordinates"].update({"file": str(coord_path.relative_to(root)), "encoding": enc, "rows": len(coords),
                                    "column_names": coords.columns.tolist(),
                                    "missing_lat_lon": int(coords[["위도", "경도"]].isna().sum().sum()),
                                    "unmatched_ridership_line_stations": missing_coords.to_dict("records")})
    dictionary.append({"dataset": "coordinates", "file": str(coord_path.relative_to(root)), "columns": coords.columns.tolist()})
    del df

    for path in sorted(root.glob("raw/02_열차혼잡도/*")):
        item = {"file": str(path.relative_to(root)), "sheets": []}
        if path.suffix == ".xlsx":
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
            for sheet in workbook:
                rows = [list(row) for row in sheet.iter_rows(values_only=True) if any(v is not None for v in row)]
                nine = "9호선" in path.name
                header_index = 1 if nine else 0
                header = rows[header_index]
                while header and header[-1] is None:
                    header.pop()
                time_count = sum(isinstance(v, str) and ":" in v and "~" in v for v in header)
                data_rows = [row for row in rows[header_index + 1:] if row[0] is not None]
                summary = {"sheet": sheet.title, "rows_with_data": len(data_rows), "columns": len(header),
                           "time_columns": time_count, "column_names": header}
                if nine:
                    summary["source_period_label"] = [v for row in rows[:2] for v in row if isinstance(v, str) and "기준일자" in v]
                else:
                    summary["line_station_count"] = len({(row[1], row[2]) for row in data_rows})
                item["sheets"].append(summary)
                dictionary.append({"dataset": "crowding", "file": str(path.relative_to(root)), "sheet": sheet.title, "columns": header})
            workbook.close()
        else:
            table, enc = csv_read(path)
            item.update({"encoding": enc, "rows": len(table), "column_names": table.columns.tolist()})
            dictionary.append({"dataset": "crowding", "file": str(path.relative_to(root)), "columns": table.columns.tolist()})
        results["crowding"].append(item)

    od_path = next(root.glob("raw/04_OD/**/지하철_역별OD_*.csv"))
    od, enc = csv_read(od_path, dtype=str)
    categories = od.iloc[:, 6:].apply(pd.to_numeric, errors="raise")
    total = pd.to_numeric(od["총_승객수"], errors="raise")
    results["od"]["station_od"] = {"file": str(od_path.relative_to(root)), "encoding": enc, "rows": len(od),
                                     "column_names": od.columns.tolist(), "dates": sorted(od["기준일자"].unique().tolist()),
                                     "has_time_of_day_column": any("시간" in col for col in od.columns),
                                     "origin_lines": sorted(od["승차_호선"].unique().tolist()),
                                     "key_duplicate_rows": int(od.duplicated(od.columns[:5].tolist()).sum()),
                                     "exact_duplicate_rows": int(od.duplicated().sum()),
                                     "passenger_sum_mismatch_rows": int((total != categories.sum(axis=1)).sum())}
    dictionary.append({"dataset": "station_od", "file": str(od_path.relative_to(root)), "columns": od.columns.tolist()})
    line_paths = sorted(root.glob("raw/04_OD/**/지하철_호선별OD_*.csv"))
    line_tables = [csv_read(p, dtype=str)[0] for p in line_paths]
    all_lines = pd.concat(line_tables, ignore_index=True)
    results["od"]["line_od"] = {"file_count": len(line_paths), "rows": len(all_lines), "dates": sorted(all_lines["기준일자"].unique().tolist()),
                                 "column_names": all_lines.columns.tolist()}
    dictionary.append({"dataset": "line_od", "file": "raw/04_OD/2026년06월_지하철/지하철/호선별OD/*.csv", "columns": all_lines.columns.tolist()})
    guide_path = next(root.glob("raw/04_OD/**/*.pdf"))
    guide = PdfReader(guide_path)
    guide_text = "\n\n".join(page.extract_text() or "" for page in guide.pages)
    (root / "metadata/OD_공식설명서_텍스트.txt").write_text(guide_text, encoding="utf-8")
    results["od"]["official_guide"] = {"file": str(guide_path.relative_to(root)), "pages": len(guide.pages)}

    results["runtime_packages"] = {x: importlib.metadata.version(x) for x in ["pandas", "openpyxl", "pypdf"]}
    (root / "metadata/data_profile.json").write_text(json.dumps(results, ensure_ascii=False, indent=2, default=int), encoding="utf-8")
    (root / "metadata/column_dictionary.json").write_text(json.dumps(dictionary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ridership": r, "coordinates": results["coordinates"], "od": {"station_rows": len(od), "line_files": len(line_paths), "station_dates": results["od"]["station_od"]["dates"]}}, ensure_ascii=False, default=int))


if __name__ == "__main__":
    main()
