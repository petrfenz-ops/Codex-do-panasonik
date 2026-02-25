param(
    [string]$ApiKey,
    [string]$ApiToken,
    [string]$BoardId,
    [string]$OutputBase = "trello_pocatecni_km",
    [switch]$IncludeOpen,
    [int]$MaxCards = 1000
)

$ErrorActionPreference = "Stop"

function Read-IfMissing {
    param(
        [string]$Value,
        [string]$Prompt
    )
    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }
    $entered = Read-Host $Prompt
    if ([string]::IsNullOrWhiteSpace($entered)) {
        throw "Chybí hodnota: $Prompt"
    }
    return $entered
}

function Invoke-TrelloGet {
    param(
        [string]$Path,
        [hashtable]$Params
    )

    $query = @{
        key = $script:ApiKey
        token = $script:ApiToken
    }

    if ($Params) {
        foreach ($k in $Params.Keys) {
            $query[$k] = $Params[$k]
        }
    }

    $base = "https://api.trello.com/1$Path"
    $qs = ($query.GetEnumerator() | ForEach-Object {
        "{0}={1}" -f [System.Web.HttpUtility]::UrlEncode([string]$_.Key), [System.Web.HttpUtility]::UrlEncode([string]$_.Value)
    }) -join "&"

    $uri = "$base`?$qs"
    return Invoke-RestMethod -Method Get -Uri $uri
}

function Parse-Date {
    param([string]$Iso)
    return [DateTimeOffset]::Parse($Iso)
}

$Regexes = @(
    [regex]"(?i)(?:p[oů]vodn[ií]|poč[aá]tečn[ií]|zač[aá]tečn[ií]|startovn[ií])?\s*(?:stav\s*)?(?:kilometr(?:y|ů)?|km)\D{0,20}(\d{1,3}(?:[ .]\d{3})+|\d{3,7})",
    [regex]"(?i)(\d{1,3}(?:[ .]\d{3})+|\d{3,7})\s*(?:km|kilometr(?:y|ů)?)"
)

function Get-KmCandidates {
    param(
        [string]$Text,
        [DateTimeOffset]$When,
        [string]$Source
    )

    $out = @()
    if ([string]::IsNullOrWhiteSpace($Text)) { return $out }

    foreach ($rx in $Regexes) {
        $matches = $rx.Matches($Text)
        foreach ($m in $matches) {
            $raw = $m.Groups[1].Value
            $digits = ($raw -replace "[^0-9]", "")
            if (-not [string]::IsNullOrWhiteSpace($digits)) {
                $out += [PSCustomObject]@{
                    Date = $When
                    Source = $Source
                    Value = [int64]$digits
                    Snippet = $m.Value
                }
            }
        }
    }
    return $out
}

function Collect-CardCandidates {
    param($Card)

    $candidates = @()

    $fallbackDate = Parse-Date $Card.dateLastActivity
    $candidates += Get-KmCandidates -Text $Card.name -When $fallbackDate -Source "Název karty (aktuální text)"
    $candidates += Get-KmCandidates -Text $Card.desc -When $fallbackDate -Source "Popis karty (aktuální text)"

    $actions = Invoke-TrelloGet -Path "/cards/$($Card.id)/actions" -Params @{
        filter = "createCard,updateCard:desc,commentCard,updateCustomFieldItem"
        fields = "type,date,data"
        limit = 1000
    }

    foreach ($action in $actions) {
        $actionDate = Parse-Date $action.date
        $type = $action.type
        $data = $action.data

        if ($type -eq "commentCard") {
            $candidates += Get-KmCandidates -Text $data.text -When $actionDate -Source "Komentář"
        }
        elseif ($type -eq "createCard") {
            if ($null -ne $data.card) {
                $candidates += Get-KmCandidates -Text $data.card.name -When $actionDate -Source "Vytvoření karty (name)"
                $candidates += Get-KmCandidates -Text $data.card.desc -When $actionDate -Source "Vytvoření karty (desc)"
            }
        }
        elseif ($type -eq "updateCard") {
            if ($null -ne $data.card) {
                $candidates += Get-KmCandidates -Text $data.card.desc -When $actionDate -Source "Úprava popisu"
            }
        }
        elseif ($type -eq "updateCustomFieldItem") {
            $v = $data.customFieldItem.value
            if ($null -ne $v) {
                if ($v.number) {
                    $candidates += Get-KmCandidates -Text ([string]$v.number) -When $actionDate -Source "Custom field (number)"
                    $digits = ([string]$v.number -replace "[^0-9]", "")
                    if ($digits) {
                        $candidates += [PSCustomObject]@{
                            Date = $actionDate
                            Source = "Custom field (number)"
                            Value = [int64]$digits
                            Snippet = [string]$v.number
                        }
                    }
                }
                if ($v.text) {
                    $candidates += Get-KmCandidates -Text ([string]$v.text) -When $actionDate -Source "Custom field (text)"
                }
            }
        }
    }

    return $candidates | Sort-Object Date
}

function Try-ConvertCsvToXlsx {
    param(
        [string]$CsvPath,
        [string]$XlsxPath
    )

    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        Write-Host "Excel není nainstalovaný, zůstane jen CSV: $CsvPath" -ForegroundColor Yellow
        return $false
    }

    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $wb = $excel.Workbooks.Open($CsvPath)
        $xlOpenXMLWorkbook = 51
        $wb.SaveAs($XlsxPath, $xlOpenXMLWorkbook)
        $wb.Close($false)
        return $true
    }
    finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

$ApiKey = Read-IfMissing -Value $ApiKey -Prompt "Zadej Trello API key"
$ApiToken = Read-IfMissing -Value $ApiToken -Prompt "Zadej Trello API token"
$BoardId = Read-IfMissing -Value $BoardId -Prompt "Zadej ID Trello boardu"

$cards = Invoke-TrelloGet -Path "/boards/$BoardId/cards" -Params @{
    fields = "id,name,desc,url,closed,dateLastActivity,idShort"
    filter = "all"
    limit = $MaxCards
}

if (-not $IncludeOpen) {
    $cards = $cards | Where-Object { $_.closed -eq $true }
}

$results = @()
$i = 0
$total = ($cards | Measure-Object).Count

foreach ($card in $cards) {
    $i++
    Write-Host "[$i/$total] Zpracovávám kartu #$($card.idShort) - $($card.name)"

    $candidates = Collect-CardCandidates -Card $card
    $first = $candidates | Select-Object -First 1

    $results += [PSCustomObject]@{
        "ID karty" = $card.idShort
        "Název karty" = $card.name
        "Archivovaná" = if ($card.closed) { "ANO" } else { "NE" }
        "Počáteční km" = if ($first) { $first.Value } else { $null }
        "Zdroj nálezu" = if ($first) { $first.Source } else { $null }
        "Datum nálezu" = if ($first) { $first.Date.ToString("yyyy-MM-dd HH:mm:ss zzz") } else { $null }
        "Ukázka textu" = if ($first) { $first.Snippet } else { $null }
        "URL" = $card.url
    }
}

$csvPath = Join-Path (Get-Location) ($OutputBase + ".csv")
$xlsxPath = Join-Path (Get-Location) ($OutputBase + ".xlsx")

$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
Write-Host "CSV export hotový: $csvPath" -ForegroundColor Green

if (Try-ConvertCsvToXlsx -CsvPath $csvPath -XlsxPath $xlsxPath) {
    Write-Host "XLSX export hotový: $xlsxPath" -ForegroundColor Green
}

$found = ($results | Where-Object { $_.'Počáteční km' -ne $null } | Measure-Object).Count
Write-Host "Hotovo. Exportováno karet: $($results.Count), nalezeno počátečních km: $found"
