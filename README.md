# Zlúčenie CSV dodacích listov

PowerShell skript na spracovanie exportovaných dodacích listov vo formáte CSV.

## Čo robí

- **Jeden súbor** — odstráni nepotrebné stĺpce
- **Viac súborov** — zlúči ich od najväčšieho po najmenší podľa počtu
  záznamov a výsledok pomenuje podľa najväčšieho z nich, potom odstráni
  stĺpce

Spracujú sa len súbory začínajúce definovanou predponou (predvolene
`zasilky-balik`).

## Použitie

1. Ulož `Zluc-Csv.ps1` a `Zluc-Csv.cmd` do rovnakého priečinka
2. V `Zluc-Csv.ps1` uprav sekciu NASTAVENIA
3. Dvojklik na `Zluc-Csv.cmd`

Priamo na `.ps1` neklikaj — Windows ho otvorí v Poznámkovom bloku,
namiesto toho, aby ho spustil.

## Nastavenia

| Premenná | Význam |
|---|---|
| `$Zdroj` | pracovný priečinok so vstupnými CSV |
| `$Vystup` | kam sa uloží výsledok |
| `$Oddelovac` | oddeľovač stĺpcov (predvolene `;`) |
| `$Predpona` | spracujú sa len súbory s touto predponou |
| `$Odstranit` | zoznam stĺpcov na odstránenie, podporuje `*` |
| `$PridatZdroj` | pridá stĺpec s názvom zdrojového súboru |
| `$RezimUpratania` | `Zmazat` / `Presunut` / `Nic` |
| `$BezPotvrdenia` | preskočí otázku pred mazaním |

`$Vystup` nechaj mimo `$Zdroj`, inak si skript pri ďalšom behu načíta
vlastný výstup.

## Kódovanie

Vstupné exporty sú v UTF-8 s BOM. Zápis je vynútený cez .NET, takže
výstup má BOM vždy — v Windows PowerShelli 5.1 aj v PowerShelli 7.
Bez BOM by Excel pokazil diakritiku.

Skript po exporte skontroluje prvé tri bajty výsledku a v súhrne
vypíše, či BOM naozaj sedí.

Samotný `Zluc-Csv.ps1` musí byť tiež uložený v UTF-8 s BOM — inak
PowerShell 5.1 zle prečíta diakritiku v názvoch stĺpcov.

## Bezpečnostné poistky

- Zdrojové súbory sa mažú až po overení, že výstup existuje a má
  správny počet záznamov
- Pred mazaním sa skript pýta na potvrdenie (`$BezPotvrdenia = $true`
  to vypne)
- Výstupný súbor nikdy nespadne do zoznamu na zmazanie
- Ak sa niektorý stĺpec zo zoznamu nenájde, skript to nahlási

`Remove-Item` obchádza kôš. Ak chceš mať zdroje po ruke, nastav
`$RezimUpratania = "Presunut"`.

## Normalizácia názvov stĺpcov

Porovnávanie hlavičiek ignoruje diakritiku, veľkosť písmen,
viacnásobné medzery a prípadný BOM. Vďaka tomu sadne `Užív. hmotnosť`
aj na `UŽÍV.  HMOTNOSŤ`.

## Office Script

V repozitári je aj `office-script.ts` — obdoba pre Excel Online, ktorá
robí to isté nad hárkami zošita. Predpona aj zoznam stĺpcov sú v oboch
riešeniach zhodné.

## Požiadavky

Windows PowerShell 5.1 alebo PowerShell 7. Skript si verziu zistí sám.
