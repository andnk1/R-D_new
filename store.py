"""
store.py - where a study's answers are kept while the preparer works through the steps.

Today: a temporary folder on the server (one small JSON file per study, plus the generated
study files). Nothing is kept long term - folders older than MAX_AGE_HOURS are deleted.

Later: swap FileStore for a database-backed store (e.g. Railway Postgres). The rest of the
app only calls get / save / folder / delete, so nothing else has to change.
"""
import json
import os
import shutil
import tempfile
import time
import uuid

MAX_AGE_HOURS = 24


class FileStore:
    def __init__(self, root=None):
        self.root = root or os.environ.get(
            "STUDY_DIR", os.path.join(tempfile.gettempdir(), "rd_feasibility_studies"))
        os.makedirs(self.root, exist_ok=True)

    # -- public API -------------------------------------------------------
    def new_id(self):
        return uuid.uuid4().hex

    def folder(self, study_id):
        path = os.path.join(self.root, _safe(study_id))
        os.makedirs(path, exist_ok=True)
        return path

    def get(self, study_id):
        path = os.path.join(self.root, _safe(study_id), "study.json")
        if not os.path.exists(path):
            return {}
        with open(path, encoding="utf-8") as f:
            return json.load(f)

    def save(self, study_id, data):
        path = os.path.join(self.folder(study_id), "study.json")
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=1)
        os.replace(tmp, path)

    def delete(self, study_id):
        shutil.rmtree(os.path.join(self.root, _safe(study_id)), ignore_errors=True)

    def prune(self):
        """Remove studies not touched in MAX_AGE_HOURS (no long-term memory)."""
        cutoff = time.time() - MAX_AGE_HOURS * 3600
        for name in os.listdir(self.root):
            p = os.path.join(self.root, name)
            try:
                if os.path.isdir(p) and os.path.getmtime(p) < cutoff:
                    shutil.rmtree(p, ignore_errors=True)
            except OSError:
                pass


def _safe(study_id):
    return "".join(c for c in str(study_id) if c.isalnum())[:64] or "none"
