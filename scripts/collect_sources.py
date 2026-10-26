"""Download official source bytes without altering their contents.

Python 3 standard library only. Run with --root /path/to/서울지하철_프로젝트.
Existing downloaded files are checked and reused unless --refresh is set.
"""
import argparse
import concurrent.futures
from datetime import datetime
import hashlib
import json
from pathlib import Path
import re
import urllib.parse
import urllib.request
import zipfile
from zoneinfo import ZoneInfo


def fetch(url, form=None):
    request = urllib.request.Request(
        url,
        data=urllib.parse.urlencode(form).encode() if form else None,
        headers={"User-Agent": "Mozilla/5.0", "Referer": "https://data.seoul.go.kr/"},
    )
    with urllib.request.urlopen(request, timeout=90) as response:
        return response.read(), dict(response.headers)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--refresh", action="store_true")
    parser.add_argument("--collection-date", default=datetime.now(ZoneInfo("Asia/Seoul")).date().isoformat())
    args = parser.parse_args()
    root = args.root.resolve()
    metadata = root / "metadata"
    metadata.mkdir(parents=True, exist_ok=True)
    source_specs = [
        ("crowding", "OA-12928", "02_열차혼잡도", ["20260630", "20260331", "20251130", "20250930"]),
        ("nine", "OA-22197", "02_열차혼잡도", ["2026년", "2025년"]),
    ]
    jobs = []
    for label, dataset_id, folder, selections in source_specs:
        page = f"https://data.seoul.go.kr/dataList/{dataset_id}/F/1/datasetView.do"
        html_bytes, _ = fetch(page)
        (metadata / f"{label}_source_page.html").write_bytes(html_bytes)
        html = html_bytes.decode("utf-8")
        matches = re.findall(r'<span title="([^"]+)" onclick="javascript:downloadFile\(\x27([^\x27]+)\x27\);">', html)
        for selection in selections:
            selected = [(name, seq) for name, seq in matches if selection in name]
            if len(selected) != 1:
                raise ValueError(f"Expected exactly one file for {dataset_id}/{selection}: {selected}")
            name, seq = selected[0]
            jobs.append({
                "dataset_id": dataset_id, "source_page": page,
                "download_url": "https://datafile.seoul.go.kr/bigfile/iot/inf/nio_download.do?&useCache=false",
                "method": "POST", "form": {"infId": dataset_id, "seqNo": "", "seq": seq, "infSeq": "1"},
                "original_filename": name, "relative_path": f"raw/{folder}/{name}",
            })

    for label, dataset_id, folder, original_name in [
        ("ridership", "15099330", "01_승하차", "서울교통공사_1_8호선 역별 일별 시간대별 승객유형별 승하차인원_20260630.csv"),
        ("coordinates", "15099316", "03_역좌표", "서울교통공사_1-8호선 역사 좌표(위경도) 정보_20250814.csv"),
    ]:
        page = f"https://www.data.go.kr/data/{dataset_id}/fileData.do"
        html_bytes, _ = fetch(page)
        (metadata / f"{label}_source_page.html").write_bytes(html_bytes)
        html = html_bytes.decode("utf-8")
        title = re.search(r"<title>(.*?)</title>", html, re.DOTALL).group(1)
        expected_date = Path(original_name).stem.rsplit("_", 1)[-1]
        if expected_date not in title:
            raise ValueError(f"Portal latest file changed; check its period before downloading: {title}")
        url = re.search(r'"contentUrl":\s*"([^"]+)"', html).group(1)
        jobs.append({"dataset_id": dataset_id, "source_page": page, "download_url": url,
                     "method": "GET", "original_filename": original_name,
                     "relative_path": f"raw/{folder}/{original_name}"})

    od_page = "https://topis.seoul.go.kr/refRoom/openRefRoom_3_4.do"
    od_endpoint = "https://topis.seoul.go.kr/refroom/selectRefRoomListASC.do"
    payload = {"blbdDivCd": "12", "bdwrDivCd": "2026", "mainBdwrRowNum": "12"}
    od_bytes, _ = fetch(od_endpoint, payload)
    (metadata / "topis_2026_file_list.json").write_bytes(od_bytes)
    od = next(x for x in json.loads(od_bytes)["rows"] if x["months"] == "06")
    jobs.append({"dataset_id": "TOPIS_OD_202606", "source_page": od_page,
                 "download_url": "https://topis.seoul.go.kr/downloadFileRefRoom.do", "method": "POST",
                 "form": {k: od[k] for k in ["apndFileNm", "apndFilePathNm", "bdwrSeq", "blbdDivCd"]},
                 "original_filename": od["apndFileNm"], "relative_path": f"raw/04_OD/{od['apndFileNm']}"})

    def download(job):
        path = root / job["relative_path"]
        path.parent.mkdir(parents=True, exist_ok=True)
        if args.refresh or not path.exists():
            data, headers = fetch(job["download_url"], job.get("form"))
            if len(data) < 50 or data.lstrip().lower().startswith((b"<!doctype", b"<html")):
                raise ValueError(f"Download returned no data: {job['original_filename']}")
            if path.suffix == ".xlsx" and not data.startswith(b"PK"):
                raise ValueError(f"Not an XLSX archive: {path}")
            partial = path.with_suffix(path.suffix + ".part")
            partial.write_bytes(data)
            partial.replace(path)
            job["content_type"] = headers.get("Content-Type", "")
        data = path.read_bytes()
        job.update({"bytes": len(data), "sha256": hashlib.sha256(data).hexdigest(), "collected_date_kst": args.collection_date})
        print(f"Saved {path.name}: {len(data):,} bytes", flush=True)
        return job

    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
        records = list(executor.map(download, jobs))
    (metadata / "download_manifest.json").write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    archive_path = root / f"raw/04_OD/{od['apndFileNm']}"
    extracted = []
    with zipfile.ZipFile(archive_path) as archive:
        for item in archive.infolist():
            if item.is_dir():
                continue
            try:
                name = item.filename.encode("cp437").decode("cp949")
            except UnicodeError:
                name = item.filename
            if not (name.startswith("지하철/") or name.endswith(".pdf")):
                continue
            path = root / "raw/04_OD/2026년06월_지하철" / name
            if not path.resolve().is_relative_to((root / "raw/04_OD/2026년06월_지하철").resolve()):
                raise ValueError(f"Unsafe ZIP path: {name}")
            path.parent.mkdir(parents=True, exist_ok=True)
            data = archive.read(item)
            path.write_bytes(data)
            extracted.append({"relative_path": str(path.relative_to(root)), "archive": str(archive_path.relative_to(root)),
                              "archive_entry": name, "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()})
    (metadata / "od_extracted_manifest.json").write_text(json.dumps(extracted, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Extracted {len(extracted)} subway/guide files; original ZIP retained.", flush=True)


if __name__ == "__main__":
    main()
