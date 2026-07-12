"""Core data models and configuration loading for the Escrow Transaction Analyst Agent."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

import yaml


@dataclass
class Category:
    code: str
    name: str
    side: str  # "inflow" | "outflow"
    keywords: list[str] = field(default_factory=list)
    description: str = ""


@dataclass
class Taxonomy:
    inflows: list[Category]
    outflows: list[Category]
    internal_transfer_keywords: list[str]

    def all(self) -> list[Category]:
        return self.inflows + self.outflows

    def by_code(self) -> dict[str, Category]:
        return {c.code: c for c in self.all()}

    def for_side(self, txn_type: str) -> list[Category]:
        """Cr transactions map to inflow categories, Dr to outflow categories."""
        return self.inflows if txn_type == "Cr" else self.outflows


@dataclass
class Transaction:
    row_no: int                    # 1-based row in the source file (audit trail)
    txn_date: date
    txn_type: str                  # "Dr" | "Cr"
    amount: float
    narration: str
    balance: Optional[float] = None
    # -- enrichment --------------------------------------------------------
    quarter: str = ""              # Q1..Q4
    category_code: str = ""        # taxonomy code, or "" if unresolved
    category_name: str = ""
    classification_source: str = ""  # RULE | AI | UNCLASSIFIED
    confidence: float = 0.0
    is_internal_transfer: bool = False
    is_duplicate: bool = False
    balance_break: bool = False
    review_reason: str = ""        # AMBIGUOUS_OR_MISSING | LOW_AI_CONFIDENCE | MULTI_MATCH
    conflict_reason: str = ""      # POTENTIALLY_INACCURATE_NARRATION detail
    waterfall_violation: str = ""  # violation type, if any

    @property
    def needs_review(self) -> bool:
        return bool(self.review_reason)

    @property
    def is_escalated_conflict(self) -> bool:
        return bool(self.conflict_reason)


def load_taxonomy(path: str | Path) -> Taxonomy:
    with open(path, encoding="utf-8") as fh:
        raw = yaml.safe_load(fh)

    def build(items, side):
        cats = []
        for it in items or []:
            cats.append(Category(
                code=str(it["code"]).strip(),
                name=str(it["name"]).strip(),
                side=side,
                keywords=[str(k).lower().strip() for k in it.get("keywords", [])],
                description=str(it.get("description", "")),
            ))
        return cats

    tax = Taxonomy(
        inflows=build(raw.get("inflow_categories"), "inflow"),
        outflows=build(raw.get("outflow_categories"), "outflow"),
        internal_transfer_keywords=[str(k).lower().strip()
                                    for k in raw.get("internal_transfer_keywords", [])],
    )
    codes = [c.code for c in tax.all()]
    dupes = {c for c in codes if codes.count(c) > 1}
    if dupes:
        raise ValueError(f"Duplicate category codes in taxonomy: {sorted(dupes)}")
    return tax


def load_settings(path: str | Path) -> dict:
    with open(path, encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def resolve_ai_enabled(settings: dict) -> bool:
    flag = settings.get("classification", {}).get("ai_enabled", "auto")
    if flag == "auto":
        return bool(os.environ.get("ANTHROPIC_API_KEY"))
    return bool(flag)
