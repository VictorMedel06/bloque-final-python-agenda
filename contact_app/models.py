from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Optional


@dataclass(frozen=True)
class Contact:
    id: Optional[int]
    name: str
    phone: str
    email: str
    created_at: Optional[datetime] = None
