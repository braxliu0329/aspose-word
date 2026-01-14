import requests
import uuid

def test_backend_indent():
    # 1. Init
    res = requests.get("http://localhost:8000/api/init")
    data = res.json()
    doc_id = data["docId"]
    version = data["version"]
    print(f"Init doc_id={doc_id}, version={version}")
    
    # 2. Insert text to have a run to select
    # Actually default doc has text. Let's find a run ID from html? 
    # It's hard to parse HTML here.
    # Instead, let's just use update_document which updates the whole doc style (and internally iterates paragraphs).
    # Or I can use update_range if I knew IDs.
    # Let's try update_document first as it also uses _apply_style_to_paragraph.
    
    client_op_id = str(uuid.uuid4())
    payload = {
        "doc_id": doc_id,
        "base_version": version,
        "client_op_id": client_op_id,
        "first_line_indent": 24.5,
        "alignment": "justify"
    }
    
    print("Sending update request:", payload)
    res = requests.post("http://localhost:8000/api/update", json=payload)
    if res.status_code != 200:
        print("Error:", res.text)
        return

    res_data = res.json()
    print("Response HTML snippet:")
    print(res_data.get("html", "")[:500])
    
    # Check if text-indent is present
    if "text-indent:24.5pt" in res_data.get("html", ""):
        print("SUCCESS: text-indent found in HTML")
    else:
        print("FAILURE: text-indent NOT found in HTML")

if __name__ == "__main__":
    try:
        test_backend_indent()
    except Exception as e:
        print(f"Test failed: {e}")
