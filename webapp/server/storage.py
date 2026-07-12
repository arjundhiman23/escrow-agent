"""Storage layer: S3 when S3_BUCKET is set (keys via env vars), local disk otherwise.

Layout (both backends):
  runs/{run_id}/meta.json
  runs/{run_id}/inputs/<files>
  runs/{run_id}/outputs/<files>
"""
import io, json, os, threading
from datetime import datetime, timezone

_LOCK = threading.Lock()


class LocalStorage:
    def __init__(self, root):
        self.root = root
        os.makedirs(root, exist_ok=True)

    def _p(self, key):
        p = os.path.normpath(os.path.join(self.root, key))
        assert p.startswith(os.path.abspath(self.root) if os.path.isabs(p) else self.root), "bad key"
        return p

    def put_bytes(self, key, data: bytes):
        p = self._p(key)
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "wb") as f:
            f.write(data)

    def get_bytes(self, key) -> bytes:
        with open(self._p(key), "rb") as f:
            return f.read()

    def exists(self, key):
        return os.path.exists(self._p(key))

    def delete_prefix(self, prefix):
        import shutil
        p = self._p(prefix)
        if os.path.isdir(p):
            shutil.rmtree(p)

    def list_dirs(self, prefix):
        p = self._p(prefix)
        if not os.path.isdir(p):
            return []
        return sorted(d for d in os.listdir(p) if os.path.isdir(os.path.join(p, d)))


class S3Storage:
    def __init__(self, bucket, region=None, prefix=""):
        import boto3
        self.s3 = boto3.client("s3", region_name=region or os.environ.get("AWS_REGION"))
        self.bucket = bucket
        self.prefix = prefix.strip("/")

    def _k(self, key):
        return f"{self.prefix}/{key}" if self.prefix else key

    def put_bytes(self, key, data: bytes):
        self.s3.put_object(Bucket=self.bucket, Key=self._k(key), Body=data)

    def get_bytes(self, key) -> bytes:
        return self.s3.get_object(Bucket=self.bucket, Key=self._k(key))["Body"].read()

    def exists(self, key):
        try:
            self.s3.head_object(Bucket=self.bucket, Key=self._k(key))
            return True
        except Exception:
            return False

    def delete_prefix(self, prefix):
        paginator = self.s3.get_paginator("list_objects_v2")
        for page in paginator.paginate(Bucket=self.bucket, Prefix=self._k(prefix)):
            objs = [{"Key": o["Key"]} for o in page.get("Contents", [])]
            if objs:
                self.s3.delete_objects(Bucket=self.bucket, Delete={"Objects": objs})

    def list_dirs(self, prefix):
        paginator = self.s3.get_paginator("list_objects_v2")
        dirs = set()
        p = self._k(prefix).rstrip("/") + "/"
        for page in paginator.paginate(Bucket=self.bucket, Prefix=p, Delimiter="/"):
            for cp in page.get("CommonPrefixes", []):
                dirs.add(cp["Prefix"][len(p):].strip("/"))
        return sorted(dirs)


def get_storage():
    bucket = os.environ.get("S3_BUCKET")
    if bucket:
        return S3Storage(bucket, os.environ.get("AWS_REGION"), os.environ.get("S3_PREFIX", "escrow-agent"))
    return LocalStorage(os.environ.get("DATA_DIR", os.path.join(os.path.dirname(__file__), "..", "..", "data")))


# ---- default template helpers (uploaded once, reused by every future run) ----
TEMPLATE_KEYS = {"catra": "templates/catra_template.xlsx", "tra": "templates/tra_template.xlsx"}


def get_template_meta(st):
    try:
        return json.loads(st.get_bytes("templates/meta.json").decode())
    except Exception:
        return {}


def set_default_template(st, kind, filename, data: bytes):
    st.put_bytes(TEMPLATE_KEYS[kind], data)
    meta = get_template_meta(st)
    meta[kind] = {"filename": filename, "uploaded_at": datetime.now(timezone.utc).isoformat(timespec="seconds")}
    st.put_bytes("templates/meta.json", json.dumps(meta).encode())
    return meta[kind]


def get_default_template(st, kind):
    """Returns (filename, bytes) or (None, None) if no default saved."""
    meta = get_template_meta(st)
    if kind not in meta or not st.exists(TEMPLATE_KEYS[kind]):
        return None, None
    return meta[kind]["filename"], st.get_bytes(TEMPLATE_KEYS[kind])


# ---- deal helpers (each deal owns its documents, extracted profile, templates, runs) ----
def read_deal(st, deal_id):
    return json.loads(st.get_bytes(f"deals/{deal_id}/meta.json").decode())


def write_deal(st, deal_id, meta):
    with _LOCK:
        st.put_bytes(f"deals/{deal_id}/meta.json", json.dumps(meta, default=str).encode())


def list_deals(st):
    return st.list_dirs("deals")


def get_deal_template_meta(st, deal_id):
    try:
        return json.loads(st.get_bytes(f"deals/{deal_id}/templates/meta.json").decode())
    except Exception:
        return {}


def set_deal_default_template(st, deal_id, kind, filename, data: bytes):
    st.put_bytes(f"deals/{deal_id}/templates/{kind}_template.xlsx", data)
    meta = get_deal_template_meta(st, deal_id)
    meta[kind] = {"filename": filename, "uploaded_at": datetime.now(timezone.utc).isoformat(timespec="seconds")}
    st.put_bytes(f"deals/{deal_id}/templates/meta.json", json.dumps(meta).encode())
    return meta[kind]


def get_deal_default_template(st, deal_id, kind):
    meta = get_deal_template_meta(st, deal_id)
    key = f"deals/{deal_id}/templates/{kind}_template.xlsx"
    if kind not in meta or not st.exists(key):
        return None, None
    return meta[kind]["filename"], st.get_bytes(key)


# ---- run meta helpers (runs now live under deals/{deal_id}/runs/{run_id}) ----
def read_run(st, deal_id, run_id):
    return json.loads(st.get_bytes(f"deals/{deal_id}/runs/{run_id}/meta.json").decode())


def write_run(st, deal_id, run_id, meta):
    with _LOCK:
        st.put_bytes(f"deals/{deal_id}/runs/{run_id}/meta.json", json.dumps(meta, default=str).encode())


def list_runs(st, deal_id):
    return st.list_dirs(f"deals/{deal_id}/runs")
