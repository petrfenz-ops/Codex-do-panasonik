# Trello export počátečního stavu kilometrů (1 soubor pro Windows)

Pro kolegyni je připravený **jeden samostatný soubor**:

- `trello_km_export_windows.ps1`

Stačí jí tenhle soubor poslat (mail / WhatsApp / SMS odkaz) a na jejím PC ho spustit v PowerShellu.
Není potřeba instalovat Python ani další knihovny.

## Co skript udělá

- připojí se na Trello API,
- projde karty boardu (standardně jen archivované),
- projde historii karet (vytvoření, komentáře, změny popisu, custom fields),
- najde **nejstarší** zmínku o stavu kilometrů,
- vytvoří výstupní tabulku:
  - vždy `CSV` (otevře se v Excelu),
  - pokud je na PC nainstalovaný Excel, vytvoří navíc i `XLSX`.

## Co kolegyně potřebuje

- Windows + PowerShell (součást Windows),
- Trello API key,
- Trello API token,
- ID boardu.

## Jak získat Trello přístupy

1. API key: `https://trello.com/power-ups/admin`
2. Token otevřením URL:

```text
https://trello.com/1/authorize?expiration=never&name=KMExport&scope=read&response_type=token&key=TVUJ_KEY
```

## Jak zjistit ID boardu

Otevřít v prohlížeči:

```text
https://api.trello.com/1/members/me/boards?key=TVUJ_KEY&token=TVUJ_TOKEN
```

A vzít `id` správného boardu.

## Spuštění (nejjednodušší)

V PowerShellu ve složce se skriptem:

```powershell
powershell -ExecutionPolicy Bypass -File .\trello_km_export_windows.ps1
```

Skript se doptá na key/token/board ID.

## Spuštění s parametry

```powershell
powershell -ExecutionPolicy Bypass -File .\trello_km_export_windows.ps1 `
  -ApiKey "TVUJ_KEY" `
  -ApiToken "TVUJ_TOKEN" `
  -BoardId "ID_BOARDU" `
  -OutputBase "vystup_km"
```

Volitelné parametry:

- `-IncludeOpen` → zahrne i nearchivované karty,
- `-MaxCards 1000` → limit počtu načtených karet.

## Poznámky

- Pokud na počítači není Excel, je to v pořádku: vznikne CSV soubor, který jde otevřít v Excelu.
- Vyhledávání km je dělané na běžné formáty (`km`, `kilometry`, `stav km`, čísla s mezerami/tečkami).
- Původní Python varianta (`trello_km_export.py`) zůstává v repozitáři, ale pro „pošlu soubor a spusť“ je doporučený PowerShell skript.
