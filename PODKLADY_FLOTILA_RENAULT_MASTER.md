# Podklady k flotile a diagnostice (Renault Master)

Nemám přímý přístup do tvého osobního prostoru ChatGPT ani do externího „Business“ workspace, takže sem dávám připravenou kostru, do které můžeme rychle doplnit konkrétní data z Fixů.

## Co sem vložit z „Fixy“

- seznam vozidel (SPZ, VIN, ročník, verze motoru)
- používaná telematika/GPS dohled (dodavatel, API/export)
- používané diagnostické nástroje (ELM/CLIP/OBD, SW verze)
- klíčové DTC chyby a jejich četnost
- mapování dat mezi GPS dohledem, Trello kartami a servisními zásahy

## Jednotný pracovní tok

1. **Sběr dat z GPS dohledu**
   - denní export: poloha, kilometry, volnoběh, spotřeba, alarmy
2. **Diagnostika Renault Master**
   - čtení DTC + freeze frame + live hodnoty
   - evidence: datum, km, stav baterie, teplota, tlak paliva, DPF hodnoty
3. **Trello synchronizace**
   - 1 karta = 1 incident / 1 vozidlo
   - povinná pole: SPZ, VIN, priorita, DTC, odhad nákladu, deadline
4. **Vyhodnocení a priorita oprav**
   - bezpečnostní závady > provozní závady > komfortní závady

## Doporučená struktura Trello karet

- **Název:** `[SPZ] Renault Master – P0/P1/P2 – stručný popis`
- **Custom fields:**
  - VIN
  - Aktuální km
  - DTC kódy
  - Stav (nové / diagnostika / objednáno / hotovo)
  - Odpovědná osoba
- **Checklist:**
  - potvrzení chyby
  - ověření živých hodnot
  - návrh opravy
  - test po opravě

## Diagnostické hodnoty, které se vyplatí pravidelně číst

- napětí baterie při startu a za běhu
- tlak paliva (požadovaná vs. skutečná hodnota)
- teplota chladicí kapaliny
- MAF/MAP odchylky
- stav DPF (soot load, regenerace)
- EGR poloha (požadovaná vs. skutečná)

## Rychlý formát incidentu

```
Vozidlo: [SPZ / VIN]
Datum/Km: [YYYY-MM-DD / km]
Symptom: [...]
DTC: [...]
Live data snapshot: [...]
Riziko provozu: [vysoké/střední/nízké]
Návrh dalšího kroku: [...]
Vazba na Trello kartu: [URL]
```

## Co udělat dál

Až sem vložíš konkrétní exporty nebo poznámky z „Fixy“, navážu tímto:

- udělám jednotný datový slovník (GPS + diagnostika + Trello),
- připravím prioritizační pravidla pro servis,
- navrhnu jednoduchý report „co opravit první“.
