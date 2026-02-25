#!/usr/bin/env python3
"""Exportuje počáteční stav kilometrů z archivovaných Trello karet do XLSX."""

from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from dataclasses import dataclass
from typing import Iterable

import requests
from openpyxl import Workbook
from openpyxl.styles import Font

API_BASE = "https://api.trello.com/1"
KM_PATTERNS = [
    re.compile(
        r"(?i)(?:p[oů]vodn[ií]|poč[aá]tečn[ií]|zač[aá]tečn[ií]|startovn[ií])?\\s*(?:stav\\s*)?(?:kilometr(?:y|ů)?|km)\\D{0,20}(\\d{1,3}(?:[ .]\\d{3})+|\\d{3,7})"
    ),
    re.compile(r"(?i)(\\d{1,3}(?:[ .]\\d{3})+|\\d{3,7})\\s*(?:km|kilometr(?:y|ů)?)"),
]


@dataclass
class Candidate:
    date: dt.datetime
    source: str
    value: int
    snippet: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Vytáhne z archivovaných Trello karet první (počáteční) stav kilometrů."
    )
    parser.add_argument("--api-key", help="Trello API key")
    parser.add_argument("--api-token", help="Trello API token")
    parser.add_argument("--board-id", help="ID boardu v Trellu")
    parser.add_argument(
        "--output",
        default="trello_pocatecni_km.xlsx",
        help="Výstupní XLSX soubor (default: trello_pocatecni_km.xlsx)",
    )
    parser.add_argument(
        "--include-open",
        action="store_true",
        help="Zahrne i nearchivované karty (jinak jen archivované).",
    )
    parser.add_argument(
        "--max-cards",
        type=int,
        default=1000,
        help="Max. počet karet načtených z boardu (default: 1000).",
    )
    return parser.parse_args()


def ask_if_missing(value: str | None, prompt: str) -> str:
    if value:
        return value
    entered = input(f"{prompt}: ").strip()
    if not entered:
        raise SystemExit(f"Chybí hodnota: {prompt}")
    return entered


def trello_get(path: str, *, key: str, token: str, params: dict | None = None) -> dict | list:
    request_params = {"key": key, "token": token}
    if params:
        request_params.update(params)
    response = requests.get(f"{API_BASE}{path}", params=request_params, timeout=30)
    response.raise_for_status()
    return response.json()


def normalize_km(raw: str) -> int:
    cleaned = re.sub(r"[^\d]", "", raw)
    return int(cleaned)


def extract_km_candidates(text: str, when: dt.datetime, source: str) -> list[Candidate]:
    candidates: list[Candidate] = []
    for pattern in KM_PATTERNS:
        for match in pattern.finditer(text):
            km = normalize_km(match.group(1))
            candidates.append(Candidate(date=when, source=source, value=km, snippet=match.group(0)))
    return candidates


def parse_iso_date(value: str) -> dt.datetime:
    return dt.datetime.fromisoformat(value.replace("Z", "+00:00"))


def collect_card_candidates(card: dict, key: str, token: str) -> list[Candidate]:
    candidates: list[Candidate] = []

    created_at = parse_iso_date(card.get("dateLastActivity"))
    if card.get("name"):
        candidates.extend(
            extract_km_candidates(card["name"], created_at, "Název karty (aktuální text)")
        )
    if card.get("desc"):
        candidates.extend(
            extract_km_candidates(card["desc"], created_at, "Popis karty (aktuální text)")
        )

    actions = trello_get(
        f"/cards/{card['id']}/actions",
        key=key,
        token=token,
        params={
            "filter": "createCard,updateCard:desc,commentCard,updateCustomFieldItem",
            "fields": "type,date,data",
            "limit": 1000,
        },
    )

    for action in actions:
        action_date = parse_iso_date(action["date"])
        action_type = action["type"]
        data = action.get("data", {})

        if action_type == "commentCard":
            text = data.get("text", "")
            candidates.extend(extract_km_candidates(text, action_date, "Komentář"))

        elif action_type == "createCard":
            card_data = data.get("card", {})
            for field_name in ("name", "desc"):
                text = card_data.get(field_name, "")
                if text:
                    candidates.extend(
                        extract_km_candidates(text, action_date, f"Vytvoření karty ({field_name})")
                    )

        elif action_type == "updateCard":
            desc = data.get("card", {}).get("desc", "")
            if desc:
                candidates.extend(extract_km_candidates(desc, action_date, "Úprava popisu"))

        elif action_type == "updateCustomFieldItem":
            custom_item = data.get("customFieldItem", {})
            value_data = custom_item.get("value", {})
            for key_name in ("number", "text"):
                if key_name in value_data and value_data[key_name]:
                    text = str(value_data[key_name])
                    candidates.extend(
                        extract_km_candidates(
                            text, action_date, f"Custom field ({key_name})"
                        )
                    )
                    if key_name == "number":
                        digits = re.sub(r"[^\d]", "", text)
                        if digits:
                            candidates.append(
                                Candidate(
                                    date=action_date,
                                    source="Custom field (number)",
                                    value=int(digits),
                                    snippet=text,
                                )
                            )

    return sorted(candidates, key=lambda c: c.date)


def iter_cards(board_id: str, key: str, token: str, max_cards: int) -> Iterable[dict]:
    cards = trello_get(
        f"/boards/{board_id}/cards",
        key=key,
        token=token,
        params={
            "fields": "id,name,desc,url,closed,dateLastActivity,idShort",
            "filter": "all",
            "limit": max_cards,
        },
    )
    if not isinstance(cards, list):
        raise RuntimeError("Neočekávaný formát odpovědi při načítání karet.")
    return cards


def export_to_xlsx(rows: list[dict], output_file: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Pocatecni_km"

    headers = [
        "ID karty",
        "Název karty",
        "Archivovaná",
        "Počáteční km",
        "Zdroj nálezu",
        "Datum nálezu",
        "Ukázka textu",
        "URL",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in rows:
        ws.append(
            [
                row["idShort"],
                row["name"],
                "ANO" if row["closed"] else "NE",
                row.get("initial_km"),
                row.get("source"),
                row.get("date"),
                row.get("snippet"),
                row["url"],
            ]
        )

    widths = [12, 48, 12, 14, 24, 24, 45, 42]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + idx)].width = width

    wb.save(output_file)


def main() -> int:
    args = parse_args()
    api_key = ask_if_missing(args.api_key, "Zadej Trello API key")
    api_token = ask_if_missing(args.api_token, "Zadej Trello API token")
    board_id = ask_if_missing(args.board_id, "Zadej ID Trello boardu")

    cards = list(iter_cards(board_id, api_key, api_token, args.max_cards))
    if not args.include_open:
        cards = [c for c in cards if c.get("closed")]

    results = []
    for index, card in enumerate(cards, start=1):
        print(f"[{index}/{len(cards)}] Zpracovávám kartu #{card['idShort']} - {card['name']}")
        candidates = collect_card_candidates(card, api_key, api_token)
        first = candidates[0] if candidates else None

        result_row = {
            "idShort": card.get("idShort"),
            "name": card.get("name"),
            "closed": card.get("closed"),
            "url": card.get("url"),
            "initial_km": first.value if first else None,
            "source": first.source if first else None,
            "date": first.date.isoformat() if first else None,
            "snippet": first.snippet if first else None,
        }
        results.append(result_row)

    export_to_xlsx(results, args.output)
    found = sum(1 for row in results if row["initial_km"] is not None)
    print(f"Hotovo. Exportováno {len(results)} karet, nalezeno {found} počátečních stavů km.")
    print(f"Soubor: {args.output}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except requests.HTTPError as exc:
        print(f"Chyba Trello API: {exc}", file=sys.stderr)
        if exc.response is not None:
            print(exc.response.text, file=sys.stderr)
        raise
