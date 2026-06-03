from fastapi.testclient import TestClient
import sys, os, json, io, zipfile, uuid
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import app

client = TestClient(app)


def test_health():
    resp = client.get("/api/health")
    assert resp.status_code == 200
    data = resp.json()
    assert data["status"] == "ok"


def test_upload_single_pdf():
    content = b"%%PDF-mock content " + str(uuid.uuid4()).encode() * 10
    resp = client.post("/api/upload", files={
        "files": ("test.pdf", content, "application/pdf")
    })
    assert resp.status_code == 200
    data = resp.json()
    assert data["files_count"] == 1
    assert len(data["files"]) == 1
    assert data["files"][0]["name"] == "test.pdf"


def test_upload_large_file():
    content = b"x" * (100 * 1024 * 1024 + 1)
    resp = client.post("/api/upload", files={
        "files": ("large.pdf", content, "application/pdf")
    })
    assert resp.status_code == 413


def test_upload_zip():
    z = io.BytesIO()
    with zipfile.ZipFile(z, "w") as zf:
        zf.writestr("doc1.pdf", b"%%PDF content 1")
        zf.writestr("doc2.pdf", b"%%PDF content 2")
    z.seek(0)
    resp = client.post("/api/upload", files={
        "files": ("files.zip", z.read(), "application/zip")
    })
    assert resp.status_code == 200
    data = resp.json()
    assert data["files_count"] == 2


def test_upload_zip_too_many():
    z = io.BytesIO()
    with zipfile.ZipFile(z, "w") as zf:
        for i in range(101):
            zf.writestr(f"doc{i}.pdf", b"%%PDF content")
    z.seek(0)
    resp = client.post("/api/upload", files={
        "files": ("files.zip", z.read(), "application/zip")
    })
    assert resp.status_code == 413


def test_process_not_found():
    resp = client.post("/api/process", json={
        "job_id": "nonexistent",
        "target_language": "pt"
    })
    assert resp.status_code == 404


def test_jobs_list():
    resp = client.get("/api/jobs")
    assert resp.status_code == 200
    assert isinstance(resp.json(), list)


def test_delete_nonexistent():
    resp = client.delete("/api/jobs/nonexistent")
    assert resp.status_code == 404


def test_outputs_nonexistent():
    resp = client.get("/api/outputs/nonexistent")
    assert resp.status_code == 404


def test_root():
    resp = client.get("/")
    assert resp.status_code == 200
    assert "PDF" in resp.text or "html" in resp.text
