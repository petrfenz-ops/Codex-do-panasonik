# Trello export počátečního stavu kilometrů

Jednoduchý skript, který z Trello boardu projde **archivované karty** a pokusí se najít
**první (nejstarší) zmínku o stavu kilometrů** v historii karty. Výsledek uloží do Excelu (`.xlsx`).

## Co umí

- načte karty z boardu (standardně jen archivované),
- načte historii akcí každé karty (vytvoření, komentáře, změny popisu, custom fields),
- najde kandidáty na hodnotu km (např. `123456 km`, `stav km: 123 456`),
- vezme časově nejstarší nález jako „počáteční stav kilometrů",
- uloží tabulku do Excelu.

## Instalace

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Kde vzít Trello API key + token

1. Key: https://trello.com/power-ups/admin
2. Token: otevři URL:
   `https://trello.com/1/authorize?expiration=never&name=KMExport&scope=read&response_type=token&key=TVUJ_KEY`

> Nahraď `TVUJ_KEY` svým API key.

## Jak zjistit ID boardu

- Otevři board v prohlížeči.
- Použij endpoint:
  `https://api.trello.com/1/members/me/boards?key=TVUJ_KEY&token=TVUJ_TOKEN`
- V odpovědi najdi správný board a zkopíruj jeho `id`.

## Spuštění

Interaktivně (skript se doptá):

```bash
python trello_km_export.py
```

Nebo rovnou s parametry:

```bash
python trello_km_export.py \
  --api-key "TVUJ_KEY" \
  --api-token "TVUJ_TOKEN" \
  --board-id "ID_BOARDU" \
  --output "vystup_km.xlsx"
```

Volitelně můžeš zahrnout i nearchivované karty:

```bash
python trello_km_export.py --include-open
```

## Výstup

Excel obsahuje sloupce:

- ID karty
- Název karty
- Archivovaná (ANO/NE)
- Počáteční km
- Zdroj nálezu
- Datum nálezu
- Ukázka textu
- URL

## Poznámky

- Snažíme se hledat formáty jako `km`, `kilometry`, `stav km`, `počáteční stav` + číslo.
- Pokud je formát na kartách hodně specifický, uprav regexy v `KM_PATTERNS`.
- Trello endpoint je omezen na počet vrácených karet (`--max-cards`, default 1000).
