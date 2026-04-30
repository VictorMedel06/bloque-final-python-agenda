from __future__ import annotations

"""
Lógica de deduplicación de contactos.

Usa RapidFuzz para calcular similitud entre nombres y agrupar contactos
potencialmente duplicados. Además, considera duplicado seguro si el teléfono
o el email coinciden exactamente.
"""

import unicodedata

from contact_app.models import Contact

try:
    from rapidfuzz import fuzz
except Exception:
    fuzz = None


class RapidFuzzNotAvailable(RuntimeError):
    pass


def is_available() -> bool:
    return fuzz is not None


def find_duplicate_groups(contacts: list[Contact], threshold: int = 88) -> list[dict]:
    if fuzz is None:
        raise RapidFuzzNotAvailable("RapidFuzz no está disponible.")

    contacts = [c for c in contacts if c.id is not None]
    if len(contacts) < 2:
        return []

    normalized = [_normalize_contact(c) for c in contacts]
    parent = {c.id: c.id for c in contacts if c.id is not None}
    best_score: dict[tuple[int, int], int] = {}

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    for i in range(len(contacts)):
        for j in range(i + 1, len(contacts)):
            a = normalized[i]
            b = normalized[j]
            score = _similarity_score(a, b)
            if score >= threshold:
                ida = int(contacts[i].id)
                idb = int(contacts[j].id)
                union(ida, idb)
                best_score[(min(ida, idb), max(ida, idb))] = score

    groups: dict[int, list[int]] = {}
    for c in contacts:
        cid = int(c.id)
        root = find(cid)
        groups.setdefault(root, []).append(cid)

    result = []
    for _, ids in groups.items():
        if len(ids) < 2:
            continue
        group_score = _group_score(ids, best_score)
        result.append({"ids": sorted(ids), "score": group_score})

    result.sort(key=lambda g: (-g["score"], len(g["ids"]), g["ids"][0]))
    return result


def merge_contacts(contacts: list[Contact]) -> Contact:
    if not contacts:
        raise ValueError("No hay contactos para fusionar.")
    contacts = [c for c in contacts if c.id is not None]
    if not contacts:
        raise ValueError("Los contactos deben tener id para fusionar.")

    primary = contacts[0]
    name = _pick_best_name([c.name for c in contacts])
    phone = _pick_first_non_empty([c.phone for c in contacts]) or primary.phone
    email = _pick_first_non_empty([c.email for c in contacts])

    return Contact(
        id=primary.id,
        name=name,
        phone=phone,
        email=email,
        created_at=primary.created_at,
    )


def _strip_accents(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def _normalize_contact(contact: Contact) -> dict[str, str]:
    name = _strip_accents(contact.name).lower().strip()
    phone = (contact.phone or "").strip()
    email = (contact.email or "").lower().strip()
    return {"name": name, "phone": phone, "email": email}


def _similarity_score(a: dict[str, str], b: dict[str, str]) -> int:
    if a["phone"] and a["phone"] == b["phone"]:
        return 100
    if a["email"] and a["email"] == b["email"]:
        return 100
    if not a["name"] or not b["name"]:
        return 0
    return int(fuzz.WRatio(a["name"], b["name"]))


def _group_score(ids: list[int], pair_scores: dict[tuple[int, int], int]) -> int:
    best = 0
    ids_sorted = sorted(ids)
    for i in range(len(ids_sorted)):
        for j in range(i + 1, len(ids_sorted)):
            key = (ids_sorted[i], ids_sorted[j])
            best = max(best, pair_scores.get(key, 0))
    return best


def _pick_first_non_empty(values: list[str]) -> str:
    for v in values:
        if v and v.strip():
            return v.strip()
    return ""


def _pick_best_name(names: list[str]) -> str:
    cleaned = [" ".join((n or "").strip().split()) for n in names]
    cleaned = [n for n in cleaned if n]
    if not cleaned:
        return ""
    cleaned.sort(key=lambda n: (len(n), n), reverse=True)
    return cleaned[0]
