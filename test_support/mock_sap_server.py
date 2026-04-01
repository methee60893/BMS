from __future__ import annotations

import argparse
import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


BASE_DIR = Path(__file__).resolve().parent
PO_SAMPLE_PATH = BASE_DIR / "mock_data" / "po_sample.json"


def load_po_samples() -> list[dict]:
    with PO_SAMPLE_PATH.open("r", encoding="utf-8") as fh:
        return json.load(fh)


class MockSapHandler(BaseHTTPRequestHandler):
    server_version = "MockSAP/1.0"

    def _send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            return None
        raw = self.rfile.read(content_length).decode("utf-8")
        return json.loads(raw)

    def do_POST(self) -> None:
        if self.path == "/ZPaymentPlan/OTBPlanUpload":
            self.handle_otb_upload()
            return
        if self.path == "/ZPaymentPlan/OTBPlanSwitch":
            self.handle_otb_switch()
            return
        self._send_json({"error": f"Unhandled POST path: {self.path}"}, status=404)

    def do_GET(self) -> None:
        if self.path.startswith("/sap/opu/odata/SAP/ZBBIK_API_2_SRV/PoSet"):
            self.handle_po_get()
            return
        self._send_json({"error": f"Unhandled GET path: {self.path}"}, status=404)

    def log_message(self, format: str, *args) -> None:
        return

    def handle_otb_upload(self) -> None:
        payload = self._read_json() or []
        results = []
        for item in payload:
            results.append(
                {
                    "version": item.get("version"),
                    "compCode": item.get("compCode"),
                    "category": item.get("category"),
                    "vendorCode": item.get("vendorCode"),
                    "segmentCode": f"({item.get('segmentCode')})" if item.get("segmentCode") else None,
                    "brandCode": item.get("brandCode"),
                    "amount": item.get("amount"),
                    "year": item.get("year"),
                    "month": item.get("month"),
                    "remark": item.get("remark"),
                    "messageType": "S",
                    "message": "Mock SAP upload success",
                }
            )
        self._send_json(
            {
                "status": {
                    "total": len(results),
                    "success": len(results),
                    "error": 0,
                    "testMode": True,
                },
                "results": results,
            }
        )

    def handle_otb_switch(self) -> None:
        payload = self._read_json() or {}
        rows = payload.get("Data", []) or []
        results = []
        for item in rows:
            results.append(
                {
                    "docyearFr": item.get("docyearFr"),
                    "periodFr": item.get("periodFr"),
                    "fmAreaFr": item.get("fmAreaFr"),
                    "catFr": item.get("catFr"),
                    "segmentFr": item.get("segmentFr"),
                    "typeFr": item.get("typeFr"),
                    "brandFr": item.get("brandFr"),
                    "vendorFr": item.get("vendorFr"),
                    "zbudget": item.get("zbudget"),
                    "docyearTo": item.get("docyearTo"),
                    "periodTo": item.get("periodTo"),
                    "fmAreaTo": item.get("fmAreaTo"),
                    "catTo": item.get("catTo"),
                    "segmentTo": item.get("segmentTo"),
                    "typeTo": item.get("typeTo"),
                    "brandTo": item.get("brandTo"),
                    "vendorTo": item.get("vendorTo"),
                    "messageType": "S",
                    "message": "Mock SAP switch success",
                }
            )
        self._send_json(
            {
                "status": {
                    "total": len(results),
                    "success": len(results),
                    "error": 0,
                    "testMode": True,
                },
                "results": results,
            }
        )

    def handle_po_get(self) -> None:
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        filter_text = (query.get("$filter", [""])[0] or "").strip()
        rows = load_po_samples()

        if "Po eq" in filter_text:
            po_no = filter_text.split("Po eq", 1)[1].strip().strip("'")
            rows = [row for row in rows if row.get("Po") == po_no]

        self._send_json({"d": {"results": rows}})


def main() -> None:
    parser = argparse.ArgumentParser(description="Mock SAP server for isolated BMS testing")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=18080)
    args = parser.parse_args()

    server = ThreadingHTTPServer((args.host, args.port), MockSapHandler)
    print(f"Mock SAP server running at http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
