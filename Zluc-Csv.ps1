# =====================================================================
#  Zlúčenie dodacích listov (CSV) a odstránenie stĺpcov
#
#  - 1 súbor  -> len odstráni stĺpce
#  - viac     -> zlúči ich od najväčšieho po najmenší (podľa počtu
#                záznamov), výsledok pomenuje podľa najväčšieho súboru,
#                potom odstráni stĺpce
#
#  Skript funguje z akéhokoľvek umiestnenia: bez parametrov spracuje
#  priečinok, v ktorom sám leží. Iný priečinok sa dá zadať parametrom
#  -Zdroj alebo pretiahnutím priečinka na Zluc-Csv.cmd.
#
#  Predpona súborov a zoznam stĺpcov sú zhodné s Office Scriptom.
# =====================================================================

param(
    # priečinok so vstupnými CSV; prázdne = priečinok, v ktorom leží skript
    [string]$Zdroj,

    # kam sa uloží výsledok; prázdne = podpriečinok "vystup" v $Zdroj
    # relatívna cesta sa berie voči $Zdroj
    [string]$Vystup
)

# ------------------------- NASTAVENIA --------------------------------
$Oddelovac = ";"                        # oddeľovač stĺpcov

# Spracujú sa len súbory začínajúce touto predponou (ako v Office Scripte)
$Predpona  = "zasilky-balik"

# Pôvodné stĺpce G, H, L, M, N — zhodné s removeHeaders v Office Scripte
# Podporuje zástupné znaky, napr. "Pozn*"
$Odstranit = @(
    "Vytvorené",
    "Podané",
    "Note",
    "Hmotnosť",
    "Užív. hmotnosť"
)

# $true = do výsledku pridá stĺpec s názvom zdrojového súboru
$PridatZdroj = $false

# $true = ak sa hlavičky súborov (po odstránení stĺpcov) líšia,
#         skript nič nezlúči a vypíše rozdiely
# $false = hlavičky zjednotí, chýbajúce bunky nechá prázdne
$PrisnaKontrola = $true

# Čo urobiť so zdrojovými súbormi po úspešnom spracovaní:
#   "Zmazat"   = natrvalo zmazať (predvolené)
#   "Presunut" = presunúť do podpriečinka _spracovane
#   "Nic"      = nechať tak
$RezimUpratania = "Zmazat"

# $true = nepýtať sa pred mazaním/presunom
$BezPotvrdenia = $false
# ---------------------------------------------------------------------


# =====================================================================
#  Kódovanie
#
#  Exporty z Balíka sú UTF-8 s BOM. Čítanie zvládne "UTF8" v oboch
#  verziách — Import-Csv si BOM odstráni sám.
#
#  Zápis je vynútený cez .NET (UTF8Encoding s BOM), lebo parameter
#  -Encoding sa v 5.1 a 7 správa rozdielne. Takto má výstup BOM vždy
#  a Excel diakritiku nepokazí.
# =====================================================================

$JeP7 = $PSVersionTable.PSVersion.Major -ge 7

$KodovanieCitanie = "UTF8"


# =====================================================================
#  Pomocné funkcie
# =====================================================================

# Zjednotí diakritiku, veľkosť písmen, viacnásobné medzery a odstráni
# prípadný BOM, aby porovnanie hlavičiek nezlyhalo na maličkosti.
function Normalizuj([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }

    $t = $text -replace "^\uFEFF", ""
    $t = ($t.Trim() -replace '\s+', ' ').Normalize([Text.NormalizationForm]::FormD)

    $sb = [Text.StringBuilder]::new()
    foreach ($z in $t.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($z) -ne 'NonSpacingMark') {
            [void]$sb.Append($z)
        }
    }
    return $sb.ToString().ToLowerInvariant()
}

# Vráti vzor z $Odstranit, ktorý na hlavičku sedí, inak $null.
function JeNaOdstranenie([string]$hlavicka) {
    $n = Normalizuj $hlavicka
    foreach ($vzor in $script:NormOdstranit) {
        if ($n -like $vzor) { return $vzor }
    }
    return $null
}

# Očistí cestu (úvodzovky, medzery, holé "C:") a urobí z nej úplnú
# cestu. Relatívna cesta sa berie voči aktuálnemu priečinku.
function UpravCestu([string]$cesta) {
    $c = $cesta.Trim().Trim('"').Trim()
    if ($c -match '^[A-Za-z]:$') { $c += '\' }
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($c)
}

$script:NormOdstranit = @($Odstranit | ForEach-Object { Normalizuj $_ })


# =====================================================================
#  Priečinky
# =====================================================================

if ([string]::IsNullOrWhiteSpace($Zdroj)) {
    $Zdroj = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
}
$Zdroj = UpravCestu $Zdroj

if (-not (Test-Path -LiteralPath $Zdroj -PathType Container)) {
    Write-Warning "Priečinok '$Zdroj' neexistuje."
    return
}

if ([string]::IsNullOrWhiteSpace($Vystup)) {
    $Vystup = Join-Path $Zdroj "vystup"
} else {
    $v = $Vystup.Trim().Trim('"').Trim()
    if (-not [IO.Path]::IsPathRooted($v)) { $v = Join-Path $Zdroj $v }
    $Vystup = UpravCestu $v
}

if ($Vystup.TrimEnd('\') -eq $Zdroj.TrimEnd('\')) {
    Write-Warning "Výstupný priečinok nesmie byť zhodný so zdrojovým — skript by pri ďalšom behu načítal vlastný výstup."
    return
}

Write-Host "`nZdroj:  $Zdroj" -ForegroundColor Cyan
Write-Host "Výstup: $Vystup" -ForegroundColor Cyan

[void][IO.Directory]::CreateDirectory($Vystup)

$subory = @(Get-ChildItem -LiteralPath $Zdroj -Filter "$Predpona*.csv" -File |
            Where-Object { $_.Extension -eq ".csv" })

if ($subory.Count -eq 0) {
    Write-Warning "V priečinku $Zdroj nie sú žiadne CSV súbory s predponou '$Predpona'."
    return
}


# =====================================================================
#  Načítanie
# =====================================================================

$nacitane = foreach ($s in $subory) {
    $data = @(Import-Csv -LiteralPath $s.FullName -Delimiter $Oddelovac -Encoding $KodovanieCitanie)

    # normalizovaný názov stĺpca -> názov stĺpca v tomto súbore
    $mapa = [ordered]@{}
    if ($data.Count -gt 0) {
        foreach ($h in $data[0].PSObject.Properties.Name) {
            $k = Normalizuj $h
            if (-not $mapa.Contains($k)) { $mapa[$k] = $h }
        }
    }

    [PSCustomObject]@{
        Nazov = $s.Name
        Pocet = $data.Count
        Data  = $data
        Mapa  = $mapa
    }
}

# od najväčšieho po najmenší; pri zhode rozhoduje názov, aby bol
# výsledok vždy rovnaký
$zoradene = @($nacitane | Sort-Object @{ Expression = "Pocet"; Descending = $true },
                                      @{ Expression = "Nazov"; Descending = $false })

Write-Host "`nNačítané súbory:" -ForegroundColor Cyan
foreach ($n in $zoradene) {
    Write-Host ("  {0,-45} {1,7} záznamov" -f $n.Nazov, $n.Pocet)
}

if ($zoradene.Count -eq 1) {
    Write-Host "  (jeden súbor — zlúčenie nie je potrebné)" -ForegroundColor DarkGray
}

$neprazdne = @($zoradene | Where-Object { $_.Pocet -gt 0 })

if ($neprazdne.Count -eq 0) {
    Write-Warning "Všetky vstupné súbory sú prázdne (bez záznamov). Nič sa nespracovalo."
    return
}


# =====================================================================
#  Stĺpce
# =====================================================================

# zjednotenie hlavičiek: normalizovaný názov -> zobrazovaný názov
# (zobrazovaný názov sa berie z najväčšieho súboru)
$vsetky = [ordered]@{}
foreach ($n in $neprazdne) {
    foreach ($k in $n.Mapa.Keys) {
        if (-not $vsetky.Contains($k)) { $vsetky[$k] = $n.Mapa[$k] }
    }
}

# vyradenie nechcených stĺpcov
$najdene  = @{}
$hlavicky = @(foreach ($k in $vsetky.Keys) {
    $vzor = JeNaOdstranenie $vsetky[$k]
    if ($vzor) { $najdene[$vzor] = $vsetky[$k] } else { $k }
})

Write-Host "`nStĺpce:" -ForegroundColor Cyan
foreach ($o in $Odstranit) {
    $n = Normalizuj $o
    if ($najdene.ContainsKey($n)) {
        Write-Host ("  odstránené: {0}" -f $najdene[$n]) -ForegroundColor DarkGray
    } else {
        Write-Warning "Stĺpec '$o' sa nenašiel — skontroluj názov alebo kódovanie súboru."
    }
}

if ($hlavicky.Count -eq 0) {
    Write-Warning "Po odstránení neostal žiadny stĺpec. Skontroluj `$Odstranit."
    return
}

Write-Host ("  ponechané:  {0}" -f (($hlavicky | ForEach-Object { $vsetky[$_] }) -join ", ")) -ForegroundColor DarkGray

# --- prísna kontrola štruktúry ---------------------------------------
if ($PrisnaKontrola -and $neprazdne.Count -gt 1) {
    $ref      = $neprazdne[0]
    $refKluce = @($ref.Mapa.Keys | Where-Object { -not (JeNaOdstranenie $ref.Mapa[$_]) })
    $hlasenie = @()

    foreach ($n in ($neprazdne | Select-Object -Skip 1)) {
        $kluce  = @($n.Mapa.Keys | Where-Object { -not (JeNaOdstranenie $n.Mapa[$_]) })
        $chyba  = @($refKluce | Where-Object { $_ -notin $kluce }    | ForEach-Object { $ref.Mapa[$_] })
        $navyse = @($kluce    | Where-Object { $_ -notin $refKluce } | ForEach-Object { $n.Mapa[$_] })

        if ($chyba.Count -gt 0 -or $navyse.Count -gt 0) {
            $hlasenie += "  $($n.Nazov)"
            if ($chyba.Count -gt 0)  { $hlasenie += "    chýba:  " + ($chyba -join ", ") }
            if ($navyse.Count -gt 0) { $hlasenie += "    navyše: " + ($navyse -join ", ") }
        }
    }

    if ($hlasenie.Count -gt 0) {
        Write-Warning "Hlavičky sa nezhodujú so súborom '$($ref.Nazov)':"
        $hlasenie | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        Write-Host "`nNič sa nezlúčilo ani nezmazalo. Ak chceš zlúčiť aj tak, nastav `$PrisnaKontrola = `$false."
        return
    }
}


# =====================================================================
#  Zlúčenie v poradí od najväčšieho
# =====================================================================

$zlucene = foreach ($n in $neprazdne) {
    foreach ($r in $n.Data) {
        $o = [ordered]@{}
        foreach ($k in $hlavicky) {
            $hodnota = ""
            if ($n.Mapa.Contains($k)) {
                $hodnota = $r.PSObject.Properties[$n.Mapa[$k]].Value
            }
            $o[$vsetky[$k]] = $hodnota
        }
        if ($PridatZdroj) { $o["ZdrojovySubor"] = $n.Nazov }
        [PSCustomObject]$o
    }
}


# =====================================================================
#  Export
#
#  Zápis ide cez .NET s vynúteným UTF-8 BOM, takže výsledok je rovnaký
#  v PowerShelli 5.1 aj 7 — nezávisle od toho, ako sa tam správa
#  parameter -Encoding.
# =====================================================================

# výstup nesie názov súboru s najväčším počtom záznamov;
# existujúci výsledok sa neprepíše
$cielovaCesta = [System.IO.Path]::GetFullPath((Join-Path $Vystup $neprazdne[0].Nazov))

if (Test-Path -LiteralPath $cielovaCesta) {
    $cas  = Get-Date -Format "yyyyMMdd-HHmmss"
    $base = [System.IO.Path]::GetFileNameWithoutExtension($cielovaCesta)
    $cielovaCesta = Join-Path $Vystup ("{0}_{1}.csv" -f $base, $cas)
    Write-Warning ("Výstup s rovnakým názvom už existuje, nový sa uloží ako {0}" -f (Split-Path $cielovaCesta -Leaf))
}

$riadky = if ($JeP7) {
    $zlucene | ConvertTo-Csv -Delimiter $Oddelovac -NoTypeInformation -UseQuotes AsNeeded
} else {
    $zlucene | ConvertTo-Csv -Delimiter $Oddelovac -NoTypeInformation
}

$utf8Bom = [System.Text.UTF8Encoding]::new($true)
[System.IO.File]::WriteAllLines($cielovaCesta, [string[]]$riadky, $utf8Bom)

# --- overenie BOM vo výstupe -----------------------------------------
$maBom = $false
if (Test-Path -LiteralPath $cielovaCesta) {
    $bajty = if ($JeP7) {
        [byte[]](Get-Content -LiteralPath $cielovaCesta -AsByteStream -TotalCount 3)
    } else {
        [byte[]](Get-Content -LiteralPath $cielovaCesta -Encoding Byte -TotalCount 3)
    }

    $maBom = ($bajty.Count -ge 3 -and $bajty[0] -eq 0xEF -and $bajty[1] -eq 0xBB -and $bajty[2] -eq 0xBF)

    if (-not $maBom) {
        Write-Warning "Výstup nemá BOM — Excel môže pokaziť diakritiku."
    }
}

Write-Host "`nHotovo." -ForegroundColor Green
Write-Host ("  Súborov:   {0}" -f $zoradene.Count)
Write-Host ("  Záznamov:  {0}" -f @($zlucene).Count)
Write-Host ("  Stĺpcov:   {0}" -f $hlavicky.Count)
Write-Host ("  Kódovanie: UTF-8 (BOM: {0})" -f $(if ($maBom) { "áno" } else { "nie" }))
Write-Host ("  Výsledok:  {0}" -f $cielovaCesta)


# =====================================================================
#  Upratanie zdrojových súborov
# =====================================================================

if ($RezimUpratania -eq "Nic") { return }

$ocakavane = @($zlucene).Count

if (-not (Test-Path -LiteralPath $cielovaCesta)) {
    Write-Warning "Výstupný súbor neexistuje. Zdroje ostávajú nedotknuté."
    return
}

$kontrola = @(Import-Csv -LiteralPath $cielovaCesta -Delimiter $Oddelovac -Encoding $KodovanieCitanie)

if ($kontrola.Count -ne $ocakavane) {
    Write-Warning ("Výstup má {0} záznamov namiesto {1}. Zdroje ostávajú nedotknuté." -f $kontrola.Count, $ocakavane)
    return
}

$naUpratanie = @($subory | Where-Object { $_.FullName -ne $cielovaCesta })

if ($naUpratanie.Count -eq 0) { return }

Write-Host "`nNa upratanie ($RezimUpratania):" -ForegroundColor Yellow
foreach ($s in $naUpratanie) { Write-Host ("  {0}" -f $s.Name) }

if (-not $BezPotvrdenia) {
    $odpoved = Read-Host "`nPokračovať? (a/n)"
    if ($odpoved -notin @("a", "A", "y", "Y")) {
        Write-Host "Zrušené, zdroje ostávajú."
        return
    }
}

if ($RezimUpratania -eq "Presunut") {
    $archiv = Join-Path $Zdroj "_spracovane"
    [void][IO.Directory]::CreateDirectory($archiv)

    foreach ($s in $naUpratanie) {
        $ciel = Join-Path $archiv $s.Name
        if (Test-Path -LiteralPath $ciel) {
            $cas  = Get-Date -Format "yyyyMMdd-HHmmss"
            $ciel = Join-Path $archiv ("{0}_{1}{2}" -f $s.BaseName, $cas, $s.Extension)
        }
        Move-Item -LiteralPath $s.FullName -Destination $ciel
    }
    Write-Host ("Presunuté do: {0}" -f $archiv) -ForegroundColor Green
}
elseif ($RezimUpratania -eq "Zmazat") {
    $naUpratanie | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force }
    Write-Host ("Zmazaných súborov: {0}" -f $naUpratanie.Count) -ForegroundColor Green
}
else {
    Write-Warning "Neznámy režim '$RezimUpratania'. Zdroje ostávajú nedotknuté."
}
