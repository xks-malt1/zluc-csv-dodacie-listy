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

Skript funguje z akéhokoľvek umiestnenia — lokálny disk, USB kľúč,
zdieľaný sieťový priečinok (aj `\\server\...`), cesty s medzerami
aj hranatými zátvorkami.

1. Ulož `Zluc-Csv.ps1` a `Zluc-Csv.cmd` do rovnakého priečinka
2. Spusti jedným zo spôsobov:
   - **dvojklik na `Zluc-Csv.cmd`** — spracuje CSV v priečinku, kde
     skript leží, výsledok uloží do jeho podpriečinka `vystup`
   - **pretiahni priečinok na `Zluc-Csv.cmd`** — spracuje CSV
     v pretiahnutom priečinku, výsledok do jeho podpriečinka `vystup`
   - **z PowerShellu:**
     ```powershell
     .\Zluc-Csv.ps1 -Zdroj "D:\Stiahnuté" -Vystup "D:\Hotové"
     ```

Priamo na `.ps1` neklikaj — Windows ho otvorí v Poznámkovom bloku,
namiesto toho, aby ho spustil.

## Parametre

| Parameter | Význam |
|---|---|
| `-Zdroj` | priečinok so vstupnými CSV; predvolene priečinok skriptu |
| `-Vystup` | kam sa uloží výsledok; predvolene `vystup` v `-Zdroj`, relatívna cesta sa berie voči `-Zdroj` |

`-Vystup` nesmie byť zhodný so `-Zdroj` — skript to odmietne, inak by
si pri ďalšom behu načítal vlastný výstup.

## Nastavenia

Sekcia NASTAVENIA v `Zluc-Csv.ps1`:

| Premenná | Význam |
|---|---|
| `$Oddelovac` | oddeľovač stĺpcov (predvolene `;`) |
| `$Predpona` | spracujú sa len súbory s touto predponou |
| `$Odstranit` | zoznam stĺpcov na odstránenie, podporuje `*` |
| `$PridatZdroj` | pridá stĺpec s názvom zdrojového súboru |
| `$PrisnaKontrola` | `$true` = pri rozdielnych hlavičkách nič nezlúči a vypíše rozdiely (predvolené) |
| `$RezimUpratania` | `Zmazat` / `Presunut` / `Nic` |
| `$BezPotvrdenia` | preskočí otázku pred mazaním |

## Kódovanie

Vstupné exporty sú v UTF-8 s BOM. Zápis je vynútený cez .NET, takže
výstup má BOM vždy — v Windows PowerShelli 5.1 aj v PowerShelli 7.
Bez BOM by Excel pokazil diakritiku.

Skript po exporte skontroluje prvé tri bajty výsledku a v súhrne
vypíše, či BOM naozaj sedí.

Samotný `Zluc-Csv.ps1` musí byť tiež uložený v UTF-8 s BOM — inak
PowerShell 5.1 zle prečíta diakritiku v názvoch stĺpcov.

## Bezpečnostné poistky

- Pri `$PrisnaKontrola = $true` sa súbory s inou štruktúrou nezlúčia —
  skript vypíše, ktoré stĺpce chýbajú alebo prebývajú, a nič nezmení
- Existujúci výsledok sa neprepíše — nový dostane časovú pečiatku
- Zdrojové súbory sa mažú až po overení, že výstup existuje a má
  správny počet záznamov
- Pred mazaním sa skript pýta na potvrdenie (`$BezPotvrdenia = $true`
  to vypne)
- Výstupný súbor nikdy nespadne do zoznamu na zmazanie
- Ak sa niektorý stĺpec zo zoznamu nenájde, skript to nahlási
- Ak sú všetky vstupné súbory prázdne, skript to povie a skončí

`Remove-Item` obchádza kôš. Ak chceš mať zdroje po ruke, nastav
`$RezimUpratania = "Presunut"`.

## Normalizácia názvov stĺpcov

Porovnávanie hlavičiek ignoruje diakritiku, veľkosť písmen,
viacnásobné medzery a prípadný BOM. Vďaka tomu sadne `Užív. hmotnosť`
aj na `UŽÍV.  HMOTNOSŤ`. Stĺpce, ktoré sa líšia len takto, sa pri
zlučovaní spoja do jedného.

## Office Script

V repozitári je aj `office-script.ts` — obdoba pre Excel Online, ktorá
robí to isté nad hárkami zošita. Predpona aj zoznam stĺpcov sú v oboch
riešeniach zhodné.

- hlavičky kontroluje ešte pred zmenou zošita — pri nezhode skončí
  chybou a zošit ostane nedotknutý
- názvy stĺpcov porovnáva s rovnakou normalizáciou ako PowerShell
- nenájdené stĺpce nahlási v konzole
- hárky pripája od najväčšieho po najmenší, po dávkach 5 000 riadkov
- menšie hárky vymaže až po overení počtu riadkov v cieľovej tabuľke

## Požiadavky

Windows PowerShell 5.1 alebo PowerShell 7. Skript si verziu zistí sám.
