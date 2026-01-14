import argparse
import json
import re
import uuid
from urllib.parse import urlencode
from urllib.request import Request, urlopen


def http_json(method: str, url: str, payload: dict | None = None) -> dict:
    data = None
    headers = {}
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = Request(url, data=data, headers=headers, method=method.upper())
    with urlopen(req, timeout=10) as resp:
        body = resp.read().decode("utf-8")
        return json.loads(body)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base-url", default="http://localhost:8000")
    parser.add_argument("--page", type=int, default=1)
    args = parser.parse_args()

    init = http_json("GET", f"{args.base_url}/api/init?{urlencode({'page': args.page})}")
    doc_id = init["docId"]
    version = init["version"]
    html = init.get("html", "")

    m = re.search(r"(Run_[0-9a-f]+)", html)
    if not m:
        raise SystemExit("No Run_ anchor found in init HTML")
    run_id = m.group(1)

    insert = {
        "doc_id": doc_id,
        "base_version": version,
        "client_op_id": str(uuid.uuid4()),
        "node_id": run_id,
        "offset": 0,
        "text": "ABCDE",
    }
    ins_res = http_json("POST", f"{args.base_url}/api/insert_text?{urlencode({'page': args.page})}", insert)
    doc_id = ins_res["docId"]
    version = ins_res["version"]

    update = {
        "doc_id": doc_id,
        "base_version": version,
        "client_op_id": str(uuid.uuid4()),
        "start_node_id": run_id,
        "start_offset": 0,
        "end_node_id": run_id,
        "end_offset": 5,
        "style": {"color": "#ff0000"},
    }
    upd_res = http_json("POST", f"{args.base_url}/api/update_range?{urlencode({'page': args.page})}", update)
    doc_id = upd_res["docId"]
    version = upd_res["version"]

    ren = http_json("GET", f"{args.base_url}/api/render?{urlencode({'page': args.page})}")
    html2 = ren.get("html", "")
    expected = f'<a name="{run_id}"><span style="color:#ff0000">ABCDE</span></a>'
    if expected not in html2:
        raise SystemExit("Regression check failed: styled inserted text not found")
    if ren.get("docId") != doc_id:
        raise SystemExit("Regression check failed: docId mismatch after render")
    if ren.get("version") != version:
        raise SystemExit("Regression check failed: version mismatch after render")

    print("OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

