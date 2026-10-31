"""Data-driven classification of *recorded feature labels*, never client content.

Rules live in feature_taxonomy.default.json and can be overridden by placing a
feature_taxonomy.json beside server.toml. A label is not proof of completion.
"""
from __future__ import annotations

import fnmatch
import json
from pathlib import Path


class TaxonomyError(ValueError):
    pass


def load_taxonomy(data_dir: Path) -> dict:
    local = data_dir / "feature_taxonomy.json"
    bundled = Path(__file__).with_name("feature_taxonomy.default.json")
    path = local if local.exists() else bundled
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as exc:
        raise TaxonomyError(f"Feature taxonomy could not be read: {path.name}") from exc
    if not isinstance(raw, dict) or not isinstance(raw.get("groups"), list) or len(raw["groups"]) > 100:
        raise TaxonomyError("Feature taxonomy must define an array of groups (max 100).")
    found = set()
    for group in raw["groups"]:
        if (not isinstance(group, dict) or not isinstance(group.get("id"), str)
                or not isinstance(group.get("label"), str)
                or not isinstance(group.get("agentic"), bool)
                or not isinstance(group.get("patterns"), list)
                or len(group["patterns"]) > 250 or not group["id"].isascii()
                or not group["id"].replace("-", "").replace("_", "").isalnum()
                or len(group["label"]) > 60 or group["id"] in found
                or any(not isinstance(pattern, str) or len(pattern) > 150 for pattern in group["patterns"])):
            raise TaxonomyError("Invalid group or pattern in feature taxonomy.")
        found.add(group["id"])
    return raw


def classify(feature: str, taxonomy: dict) -> dict:
    for group in taxonomy["groups"]:
        if any(fnmatch.fnmatchcase(feature.casefold(), pattern.casefold()) for pattern in group["patterns"]):
            return {"id": group["id"], "label": group["label"], "agentic": group["agentic"]}
    return {"id": "other", "label": "Other recorded functions", "agentic": False}
