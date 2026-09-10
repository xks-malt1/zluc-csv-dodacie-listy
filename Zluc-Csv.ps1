# =====================================================================
#  Zlúčenie dodacích listov (CSV) a odstránenie stĺpcov
#
#  - 1 súbor  -> len odstráni stĺpce
#  - viac     -> zlúči ich od najväčšieho po najmenší (podľa počtu
#                záznamov), výsledok pomenuje podľa najväčšieho súboru,
#                potom odstráni stĺpce
#
#  Predpona súborov a zoznam stĺpcov sú zhodné s Office Scriptom.
# =====================================================================

# ------------------------- NASTAVENIA --------------------------------
$Zdroj     = "C:\data"                  # pracovný priečinok so vstupnými CSV
$Vystup    = "C:\data\vystup"           # kam sa uloží výsledok
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


# --- normalizácia názvov stĺpcov -------------------------------------
#  Zjednotí diakritiku, veľkosť písmen, viacnásobné medzery a odstráni
#  prípadný BOM, aby porovnanie hlavičiek nezlyhalo na maličkosti.
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


if (-not (Test-Path $Vystup)) {
    New-Item -ItemType Directory -Path $Vystup | Out-Null
}

$subory = @(Get-ChildItem -Path $Zdroj -Filter "$Predpona*.csv" -File)

if ($subory.Count -eq 0) {
    Write-Warning "V priečinku $Zdroj nie sú žiadne CSV súbory s predponou '$Predpona'."
    return
}

# --- načítanie súborov a spočítanie záznamov -------------------------
$nacitane = foreach ($s in $subory) {
    $data = @(Import-Csv -Path $s.FullName -Delimiter $Oddelovac -Encoding $KodovanieCitanie)
    [PSCustomObject]@{
        Nazov = $s.Name
        Pocet = $data.Count
        Data  = $data
    }
}

# --- zoradenie od najväčšieho po najmenší ----------------------------
$zoradene = @($nacitane | Sort-Object Pocet -Descending)

Write-Host "`nNačítané súbory:" -ForegroundColor Cyan
foreach ($n in $zoradene) {
    Write-Host ("  {0,-45} {1,7} záznamov" -f $n.Nazov, $n.Pocet)
}

if ($zoradene.Count -eq 1) {
    Write-Host "  (jeden súbor — zlúčenie nie je potrebné)" -ForegroundColor DarkGray
}

# výstup nesie názov súboru s najväčším počtom záznamov
$cielovaCesta = Join-Path $Vystup $zoradene[0].Nazov

# --- zjednotenie hlavičiek zo všetkých súborov -----------------------
$vsetky = [System.Collections.Specialized.OrderedDictionary]::new()
foreach ($n in $zoradene) {
    if ($n.Pocet -gt 0) {
        foreach ($h in $n.Data[0].PSObject.Properties.Name) {
            if (-not $vsetky.Contains($h)) { $vsetky.Add($h, $true) }
        }
    }
}

# --- vyradenie nechcených stĺpcov ------------------------------------
$normOdstranit = $Odstranit | ForEach-Object { Normalizuj $_ }
$najdene = @{}

$hlavicky = @($vsetky.Keys) | Where-Object {
    $norm  = Normalizuj $_
    $zhoda = @($normOdstranit | Where-Object { $norm -like $_ })
    if ($zhoda.Count -gt 0) { $najdene[$zhoda[0]] = $_ }
    $zhoda.Count -eq 0
}

Write-Host "`nStĺpce:" -ForegroundColor Cyan
foreach ($o in $Odstranit) {
    $n = Normalizuj $o
    if ($najdene.ContainsKey($n)) {
        Write-Host ("  odstránené: {0}" -f $najdene[$n]) -ForegroundColor DarkGray
    } else {
        Write-Warning "Stĺpec '$o' sa nenašiel — skontroluj názov alebo kódovanie súboru."
    }
}
Write-Host ("  ponechané:  {0}" -f ($hlavicky -join ", ")) -ForegroundColor DarkGray

if ($hlavicky.Count -eq 0) {
    Write-Warning "Po odstránení neostal žiadny stĺpec. Skontroluj `$Odstranit."
    return
}

# --- zlúčenie v poradí od najväčšieho --------------------------------
$zlucene = foreach ($n in $zoradene) {
    if ($PridatZdroj) {
        $n.Data | Select-Object ($hlavicky + @{ Name = "ZdrojovySubor"; Expression = { $n.Nazov } })
    } else {
        $n.Data | Select-Object $hlavicky
    }
}

# --- export ----------------------------------------------------------
#  Zápis ide cez .NET s vynúteným UTF-8 BOM, takže výsledok je rovnaký
#  v PowerShelli 5.1 aj 7 — nezávisle od toho, ako sa tam správa
#  parameter -Encoding.
$cielovaCesta = [System.IO.Path]::GetFullPath($cielovaCesta)

$riadky = if ($JeP7) {
    $zlucene | ConvertTo-Csv -Delimiter $Oddelovac -NoTypeInformation -UseQuotes AsNeeded
} else {
    $zlucene | ConvertTo-Csv -Delimiter $Oddelovac -NoTypeInformation
}

$utf8Bom = [System.Text.UTF8Encoding]::new($true)
[System.IO.File]::WriteAllLines($cielovaCesta, $riadky, $utf8Bom)

# --- overenie BOM vo výstupe -----------------------------------------
$maBom = $false
if (Test-Path $cielovaCesta) {
    $bajty = if ($JeP7) {
        [byte[]](Get-Content $cielovaCesta -AsByteStream -TotalCount 3)
    } else {
        [byte[]](Get-Content $cielovaCesta -Encoding Byte -TotalCount 3)
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

if (-not (Test-Path $cielovaCesta)) {
    Write-Warning "Výstupný súbor neexistuje. Zdroje ostávajú nedotknuté."
    return
}

$kontrola = @(Import-Csv -Path $cielovaCesta -Delimiter $Oddelovac -Encoding $KodovanieCitanie)

if ($kontrola.Count -ne $ocakavane) {
    Write-Warning ("Výstup má {0} záznamov namiesto {1}. Zdroje ostávajú nedotknuté." -f $kontrola.Count, $ocakavane)
    return
}

$naUpratanie = @($subory | Where-Object { $_.FullName -ne (Resolve-Path $cielovaCesta).Path })

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
    if (-not (Test-Path $archiv)) { New-Item -ItemType Directory -Path $archiv | Out-Null }

    foreach ($s in $naUpratanie) {
        $ciel = Join-Path $archiv $s.Name
        if (Test-Path $ciel) {
            $cas  = Get-Date -Format "yyyyMMdd-HHmmss"
            $ciel = Join-Path $archiv ("{0}_{1}{2}" -f $s.BaseName, $cas, $s.Extension)
        }
        Move-Item -Path $s.FullName -Destination $ciel
    }
    Write-Host ("Presunuté do: {0}" -f $archiv) -ForegroundColor Green
}
elseif ($RezimUpratania -eq "Zmazat") {
    $naUpratanie | Remove-Item -Force
    Write-Host ("Zmazaných súborov: {0}" -f $naUpratanie.Count) -ForegroundColor Green
}
else {
    Write-Warning "Neznámy režim '$RezimUpratania'. Zdroje ostávajú nedotknuté."
}
