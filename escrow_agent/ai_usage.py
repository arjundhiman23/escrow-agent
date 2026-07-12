"""Token usage + cost tracking for Anthropic API calls made by the agent.

Every call site that hits the Anthropic API should call `record()` with the
raw `usage` object from the response. Rates are per-million-tokens (USD) and
are intentionally explicit/editable here rather than hidden in call sites, so
the cost KPI on the frontend stays accurate as pricing changes.
"""
from dataclasses import dataclass, field, asdict

# USD per 1M tokens: (input, output). Update if Anthropic pricing changes.
# Sonnet 5 introductory rate runs through 2026-08-31; standard $3/$15 after.
PRICING = {
    "claude-haiku-4-5-20251001": {"input": 1.00, "output": 5.00},
    "claude-sonnet-5":           {"input": 2.00, "output": 10.00},   # introductory, through 2026-08-31
    "claude-opus-4-8":           {"input": 5.00, "output": 25.00},
    "claude-sonnet-4-6":         {"input": 3.00, "output": 15.00},   # legacy, kept for old configs
}
DEFAULT_RATE = {"input": 3.00, "output": 15.00}  # fallback if model string is unrecognised


def rate_for(model: str):
    return PRICING.get(model, DEFAULT_RATE)


def cost_usd(model: str, input_tokens: int, output_tokens: int) -> float:
    r = rate_for(model)
    return round(input_tokens / 1e6 * r["input"] + output_tokens / 1e6 * r["output"], 6)


@dataclass
class UsageEvent:
    purpose: str          # "classification" | "sanction_extraction"
    model: str
    input_tokens: int
    output_tokens: int
    cost_usd: float


class UsageTracker:
    """Accumulates events during a single pipeline run."""
    def __init__(self):
        self.events: list[UsageEvent] = []

    def record(self, purpose: str, model: str, usage) -> UsageEvent:
        # `usage` is the Anthropic SDK's Usage object (or any obj/dict with
        # input_tokens/output_tokens) attached to resp.usage.
        it = getattr(usage, "input_tokens", None)
        ot = getattr(usage, "output_tokens", None)
        if it is None and isinstance(usage, dict):
            it, ot = usage.get("input_tokens", 0), usage.get("output_tokens", 0)
        it, ot = int(it or 0), int(ot or 0)
        ev = UsageEvent(purpose=purpose, model=model, input_tokens=it, output_tokens=ot,
                        cost_usd=cost_usd(model, it, ot))
        self.events.append(ev)
        return ev

    def summary(self) -> dict:
        total_in = sum(e.input_tokens for e in self.events)
        total_out = sum(e.output_tokens for e in self.events)
        total_cost = round(sum(e.cost_usd for e in self.events), 6)
        by_purpose = {}
        for e in self.events:
            p = by_purpose.setdefault(e.purpose, {"input_tokens": 0, "output_tokens": 0, "cost_usd": 0.0, "calls": 0})
            p["input_tokens"] += e.input_tokens
            p["output_tokens"] += e.output_tokens
            p["cost_usd"] = round(p["cost_usd"] + e.cost_usd, 6)
            p["calls"] += 1
        return {
            "input_tokens": total_in, "output_tokens": total_out, "cost_usd": total_cost,
            "calls": len(self.events), "by_purpose": by_purpose,
            "events": [asdict(e) for e in self.events],
        }
