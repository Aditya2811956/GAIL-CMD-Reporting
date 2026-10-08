"""
migrate_collection.py

Copies all chunks from the local ChromaDB collection (cmd_reports_v2, v1.5.5)
to the remote ChromaDB server (cmd_reports_2, v0.6.3) using raw HTTP requests.

Usage:
    python migrate_collection.py [--dry-run]

    --dry-run   Print counts and sample metadata without writing to remote.
"""

import argparse
import sys
import time
from pathlib import Path

import chromadb
import requests

# ── config ────────────────────────────────────────────────────────────────────
LOCAL_HOST       = "127.0.0.1"
LOCAL_PORT       = 8001
LOCAL_COLLECTION = "cmd_reports_v2"

REMOTE_HOST       = "52.172.103.119"
REMOTE_PORT       = 6062
REMOTE_COLLECTION = "cmd_reports_v2"

BATCH_SIZE = 50   # keep small — embeddings are large


# ── remote helpers (v0.6.3 REST API) ─────────────────────────────────────────

def remote_url(path: str) -> str:
    return f"http://{REMOTE_HOST}:{REMOTE_PORT}/api/v1/{path.lstrip('/')}"


def get_or_create_remote_collection(name: str) -> str:
    """Create collection on remote if it doesn't exist. Returns collection id."""
    resp = requests.post(
        remote_url("collections"),
        json={"name": name, "get_or_create": True},
        timeout=30,
    )
    resp.raise_for_status()
    col_id = resp.json()["id"]
    print(f"  Remote collection '{name}' ready (id={col_id})")
    return col_id


def remote_collection_count(col_id: str) -> int:
    resp = requests.get(remote_url(f"collections/{col_id}/count"), timeout=10)
    resp.raise_for_status()
    return resp.json()


def add_to_remote(col_id: str, ids, embeddings, documents, metadatas):
    # Convert numpy arrays to plain Python lists for JSON serialisation
    embeddings_list = [e.tolist() if hasattr(e, "tolist") else list(e) for e in embeddings]
    resp = requests.post(
        remote_url(f"collections/{col_id}/upsert"),  # upsert handles existing IDs gracefully
        json={
            "ids":        ids,
            "embeddings": embeddings_list,
            "documents":  documents,
            "metadatas":  metadatas,
        },
        timeout=120,
    )
    if resp.status_code not in (200, 201):
        print(f"  [WARN] Remote upsert returned {resp.status_code}: {resp.text[:200]}")
        return False
    return True


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Read only, do not write to remote")
    args = parser.parse_args()

    # ── connect to local collection ───────────────────────────────────────────
    print(f"\nConnecting to local ChromaDB ({LOCAL_HOST}:{LOCAL_PORT})...")
    local_client = chromadb.HttpClient(host=LOCAL_HOST, port=LOCAL_PORT)
    local_col = local_client.get_collection(LOCAL_COLLECTION)
    total = local_col.count()
    print(f"  Local collection '{LOCAL_COLLECTION}': {total} chunks")

    if total == 0:
        print("  Nothing to migrate.")
        sys.exit(0)

    # ── create remote collection ──────────────────────────────────────────────
    if not args.dry_run:
        print(f"\nConnecting to remote ChromaDB ({REMOTE_HOST}:{REMOTE_PORT})...")
        col_id = get_or_create_remote_collection(REMOTE_COLLECTION)
        existing = remote_collection_count(col_id)
        print(f"  Remote collection already has {existing} chunks")
    else:
        print("\n[DRY RUN] Skipping remote connection.")

    # ── paginate and migrate ──────────────────────────────────────────────────
    print(f"\nMigrating {total} chunks in batches of {BATCH_SIZE}...")
    offset   = 0
    migrated = 0
    skipped  = 0

    while offset < total:
        batch = local_col.get(
            limit=BATCH_SIZE,
            offset=offset,
            include=["embeddings", "documents", "metadatas"],
        )

        ids        = batch["ids"]
        embeddings = batch["embeddings"]
        documents  = batch["documents"]
        metadatas  = batch["metadatas"]

        if not ids:
            break

        if args.dry_run:
            print(f"  [DRY RUN] Would upload batch {offset}–{offset + len(ids) - 1} ({len(ids)} chunks)")
            if offset == 0:
                print(f"  Sample metadata: {metadatas[0]}")
        else:
            ok = add_to_remote(col_id, ids, embeddings, documents, metadatas)
            if ok:
                migrated += len(ids)
            else:
                skipped += len(ids)

            pct = int((offset + len(ids)) / total * 100)
            print(f"  [{pct:3d}%] Uploaded {offset + len(ids)}/{total} chunks", end="\r")

        offset += len(ids)
        time.sleep(0.1)   # be gentle on the remote

    if not args.dry_run:
        print()   # newline after \r
        final_count = remote_collection_count(col_id)
        print(f"\nDone. Migrated: {migrated}  Skipped: {skipped}")
        print(f"Remote collection now has {final_count} chunks.")
    else:
        print(f"\n[DRY RUN] Would migrate {total} chunks total.")


if __name__ == "__main__":
    main()
