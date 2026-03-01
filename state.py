"""
Bijhoudt welke NMBS-bestellingen al verwerkt zijn, zodat er nooit dubbele
rijen in de onkostennota terechtkomen.
"""
import json
from pathlib import Path


def load_state(state_file: Path) -> dict:
    """Laad de verwerkte bestellingen uit het state-bestand."""
    empty: dict = {"processed": [], "skipped_weekend": [], "metadata": {}}
    if state_file.exists():
        try:
            with open(state_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError) as exc:
            print(f"  Waarschuwing: state-bestand onleesbaar ({exc}), start met lege state.")
            return empty
    return empty


def save_state(state: dict, state_file: Path) -> None:
    """Sla de huidige state op."""
    with open(state_file, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def is_processed(order_number: str, state: dict) -> bool:
    return order_number in state.get("processed", [])


def is_skipped(order_number: str, state: dict) -> bool:
    return order_number in state.get("skipped_weekend", [])


def mark_processed(
    order_number: str, state: dict, metadata: dict | None = None
) -> None:
    if order_number not in state["processed"]:
        state["processed"].append(order_number)
    if metadata is not None:
        state.setdefault("metadata", {})[order_number] = metadata


def get_metadata(order_number: str, state: dict) -> dict | None:
    """Geeft de opgeslagen Excel-metadata voor een bestelling, of None."""
    return state.get("metadata", {}).get(order_number)


def mark_skipped_weekend(order_number: str, state: dict) -> None:
    """Markeer een weekend/feestdag-ticket als permanent overgeslagen."""
    if order_number not in state.setdefault("skipped_weekend", []):
        state["skipped_weekend"].append(order_number)
