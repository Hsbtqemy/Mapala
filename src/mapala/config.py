"""Configuration minimale pour Mapala."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


class LaConcordeError(Exception):
    """Exception de base (compatibilité)."""


class ConfigError(LaConcordeError, ValueError):
    """Erreur de validation de la configuration."""


@dataclass
class ConcatSource:
    """Source pour une concaténation (colonne + préfixe optionnel)."""

    col: str
    prefix: str = ""

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "ConcatSource":
        col = str(d.get("col", "")).strip()
        if not col:
            raise ConfigError("concat source 'col' requis")
        return cls(col=col, prefix=str(d.get("prefix", "")))
