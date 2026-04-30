from __future__ import annotations

import re
from typing import Optional


_EMAIL_RE = re.compile(
    r"^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$",
    re.IGNORECASE,
)


def normalize_phone(raw: Optional[str]) -> str:
    if raw is None:
        return ""
    cleaned = raw.strip()
    if not cleaned:
        return ""
    cleaned = re.sub(r"[\s\-()]+", "", cleaned)
    return cleaned


def validate_name(name: Optional[str]) -> str:
    if name is None:
        raise ValueError("El nombre es obligatorio.")
    cleaned = " ".join(name.strip().split())
    if not cleaned:
        raise ValueError("El nombre es obligatorio.")
    for ch in cleaned:
        if ch.isalpha() or ch in {" ", "-", "’", "'"}:
            continue
        raise ValueError("El nombre solo puede contener letras, espacios y guiones.")
    return cleaned


def validate_phone(phone: Optional[str]) -> str:
    cleaned = normalize_phone(phone)
    if not cleaned:
        raise ValueError("El teléfono es obligatorio.")
    if cleaned.startswith("+"):
        digits = cleaned[1:]
    else:
        digits = cleaned
    if not digits.isdigit():
        raise ValueError("El teléfono solo puede contener números (opcionalmente + al inicio).")
    if not (7 <= len(digits) <= 15):
        raise ValueError("El teléfono debe tener entre 7 y 15 dígitos.")
    return cleaned


def validate_email(email: Optional[str]) -> str:
    if email is None:
        return ""
    cleaned = email.strip()
    if not cleaned:
        return ""
    if not _EMAIL_RE.match(cleaned):
        raise ValueError("El correo electrónico no tiene un formato válido.")
    return cleaned.lower()
