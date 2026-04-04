[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Notebook,

    [Parameter(Mandatory = $true)]
    [string[]]$Sections,

    [Parameter(Mandatory = $true)]
    [string]$Output,

    [int]$LimitPages = 0,

    [switch]$ListOnly,

    [switch]$StopOnError
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$HierarchyScopePages = 4
$XmlSchema2013 = 2
$PublishFormatWord = 5

function Format-HResult {
    param([System.Exception]$Exception)

    if ($null -eq $Exception) {
        return "0x00000000"
    }
    return ("0x{0:X8}" -f ($Exception.HResult -band 0xFFFFFFFF))
}

function Normalize-ForCompare {
    param([string]$Value)

    if ($null -eq $Value) {
        return ""
    }

    $normalized = $Value.Trim().ToLowerInvariant()
    $normalized = [regex]::Replace($normalized, "[‐-―−]", "-")
    $normalized = [regex]::Replace($normalized, "\s*-\s*", "-")
    $normalized = [regex]::Replace($normalized, "\s+", " ")
    return $normalized
}

function ConvertTo-SafeName {
    param(
        [string]$Value,
        [string]$Fallback = "untitled"
    )

    if ($null -eq $Value) {
        $safe = ""
    }
    else {
        $safe = $Value.Trim()
    }
    $safe = [regex]::Replace($safe, '[\\/:*?"<>|]+', "-")
    $safe = [regex]::Replace($safe, '\s+', ' ')
    $safe = $safe.Trim(' ', '.')

    if ([string]::IsNullOrWhiteSpace($safe)) {
        return $Fallback
    }
    return $safe
}

function Get-SectionPath {
    param([System.Xml.XmlNode]$SectionNode)

    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add($SectionNode.Attributes["name"].Value)

    $parent = $SectionNode.ParentNode
    while ($null -ne $parent) {
        if ($parent.LocalName -eq "SectionGroup") {
            $parts.Insert(0, $parent.Attributes["name"].Value)
        }
        elseif ($parent.LocalName -eq "Notebook") {
            $parts.Insert(0, $parent.Attributes["name"].Value)
            break
        }
        $parent = $parent.ParentNode
    }

    return ($parts -join "/")
}

function Parse-PageDate {
    param([System.Xml.XmlNode]$PageNode)

    $raw = ""
    $dateAttr = $PageNode.Attributes["dateTime"]
    if ($null -ne $dateAttr) {
        $raw = $dateAttr.Value
    }
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $null
    }

    try {
        return [datetime]::Parse(
            $raw,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind
        )
    }
    catch {
        return $null
    }
}

function Get-UniqueDocxPath {
    param(
        [string]$Directory,
        [string]$BaseName
    )

    $candidate = Join-Path $Directory ($BaseName + ".docx")
    if (-not (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }

    $counter = 2
    while ($true) {
        $nextName = "{0}-{1}" -f $BaseName, $counter
        $candidate = Join-Path $Directory ($nextName + ".docx")
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
        $counter++
    }
}

try {
    $apartmentState = [System.Threading.Thread]::CurrentThread.ApartmentState
    if ($apartmentState -ne [System.Threading.ApartmentState]::STA) {
        Write-Warning ((
            "Aktueller Thread ist {0}. OneNote COM laeuft stabiler in STA. " +
            "Empfehlung: powershell.exe -STA -ExecutionPolicy Bypass -File .\export_onenote_com.ps1 ..."
        ) -f $apartmentState)
    }

    Write-Host "Initialisiere OneNote COM..."
    try {
        $one = New-Object -ComObject OneNote.Application
    }
    catch {
        $hr = Format-HResult $_.Exception
        throw ((
            "OneNote COM konnte nicht initialisiert werden (HRESULT {0}). " +
            "Pruefe: OneNote Desktop (M365/2016) installiert, nicht nur OneNote for Windows 10."
        ) -f $hr)
    }

    $hierarchyXml = ""
    try {
        $one.GetHierarchy("", $HierarchyScopePages, [ref]$hierarchyXml, $XmlSchema2013)
    }
    catch {
        $hr = Format-HResult $_.Exception
        throw ((
            "GetHierarchy von OneNote COM fehlgeschlagen (HRESULT {0}). " +
            "Oeffne OneNote Desktop einmal manuell und lass die Notizbuecher komplett synchronisieren."
        ) -f $hr)
    }

    [xml]$doc = $hierarchyXml
    $nsUri = $doc.DocumentElement.NamespaceURI
    $nsmgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsmgr.AddNamespace("one", $nsUri)

    $allNotebooks = @($doc.SelectNodes("//one:Notebook", $nsmgr))
    $wantedNotebook = Normalize-ForCompare $Notebook
    $notebookNode = $allNotebooks | Where-Object {
        (Normalize-ForCompare $_.Attributes["name"].Value) -eq $wantedNotebook
    } | Select-Object -First 1

    if ($null -eq $notebookNode) {
        $available = $allNotebooks | ForEach-Object { $_.Attributes["name"].Value } | Sort-Object
        throw "Notebook '$Notebook' nicht gefunden. Verfuegbar: $($available -join ', ')"
    }

    $sectionNodes = @($notebookNode.SelectNodes(".//one:Section", $nsmgr))
    $targetSections = New-Object System.Collections.Generic.List[System.Xml.XmlNode]

    $requestedSections = New-Object System.Collections.Generic.List[string]
    foreach ($rawSection in $Sections) {
        if ([string]::IsNullOrWhiteSpace($rawSection)) {
            continue
        }

        $splitItems = @($rawSection -split ",")
        foreach ($item in $splitItems) {
            $clean = $item.Trim().Trim("'", '"')
            if (-not [string]::IsNullOrWhiteSpace($clean)) {
                $requestedSections.Add($clean)
            }
        }
    }

    if ($requestedSections.Count -eq 0) {
        throw "Keine gueltigen Abschnittsnamen in -Sections uebergeben."
    }

    foreach ($sectionName in $requestedSections) {
        $wantedSection = Normalize-ForCompare $sectionName
        $hits = @($sectionNodes | Where-Object {
            $nodeName = Normalize-ForCompare $_.Attributes["name"].Value
            $nodePath = Normalize-ForCompare (Get-SectionPath $_)
            $nodeName -eq $wantedSection -or $nodePath -eq $wantedSection
        })

        if ($hits.Count -eq 0) {
            $availableSections = $sectionNodes | ForEach-Object { Get-SectionPath $_ } | Sort-Object
            throw "Abschnitt '$sectionName' nicht gefunden. Verfuegbar: $($availableSections -join ', ')"
        }

        if ($hits.Count -gt 1) {
            $paths = $hits | ForEach-Object { Get-SectionPath $_ }
            throw "Abschnitt '$sectionName' ist mehrfach vorhanden: $($paths -join ', ')"
        }

        $targetSections.Add($hits[0])
    }

    $outRoot = (Resolve-Path -LiteralPath (New-Item -ItemType Directory -Path $Output -Force).FullName).Path
    $notebookDir = Join-Path $outRoot (ConvertTo-SafeName -Value $notebookNode.Attributes["name"].Value -Fallback "notebook")
    New-Item -ItemType Directory -Path $notebookDir -Force | Out-Null

    $report = New-Object System.Collections.Generic.List[object]
    $totalPages = 0
    $totalExported = 0
    $totalFailed = 0

    Write-Host ("Notebook: {0}" -f $notebookNode.Attributes["name"].Value)

    foreach ($sectionNode in $targetSections) {
        $sectionName = $sectionNode.Attributes["name"].Value
        $sectionPath = Get-SectionPath $sectionNode
        $sectionDir = Join-Path $notebookDir (ConvertTo-SafeName -Value $sectionName -Fallback "section")
        New-Item -ItemType Directory -Path $sectionDir -Force | Out-Null

        $pages = @($sectionNode.SelectNodes(".//one:Page", $nsmgr))
        $pages = $pages | Sort-Object {
            $parsed = Parse-PageDate $_
            if ($null -eq $parsed) {
                return [datetime]::MaxValue
            }
            return $parsed
        }

        if ($LimitPages -gt 0 -and $pages.Count -gt $LimitPages) {
            $pages = @($pages[0..($LimitPages - 1)])
        }

        Write-Host ""
        Write-Host ("Abschnitt: {0}" -f $sectionPath)

        $sectionExported = 0
        $sectionFailed = 0
        for ($i = 0; $i -lt $pages.Count; $i++) {
            $page = $pages[$i]
            $pageId = $page.Attributes["ID"].Value

            $title = ""
            $titleAttr = $page.Attributes["name"]
            if ($null -ne $titleAttr) {
                $title = $titleAttr.Value
            }
            if ([string]::IsNullOrWhiteSpace($title)) {
                $title = "page-{0:D3}" -f ($i + 1)
            }

            $pageDateTimeRaw = ""
            $dateTimeAttr = $page.Attributes["dateTime"]
            if ($null -ne $dateTimeAttr) {
                $pageDateTimeRaw = $dateTimeAttr.Value
            }

            $dt = Parse-PageDate $page
            $prefix = if ($null -ne $dt) { $dt.ToString("yyyy-MM-dd_HH-mm") } else { "undated" }
            $baseName = ConvertTo-SafeName -Value ("{0} {1}" -f $prefix, $title) -Fallback ("page-{0:D3}" -f ($i + 1))
            $docxPath = Get-UniqueDocxPath -Directory $sectionDir -BaseName $baseName

            Write-Host ("  -> Seite {0}/{1}: {2}" -f ($i + 1), $pages.Count, $title)

            $entry = [ordered]@{
                section = $sectionName
                sectionPath = $sectionPath
                pageIndex = ($i + 1)
                pageId = $pageId
                title = $title
                dateTime = $pageDateTimeRaw
                output = $docxPath
                status = "pending"
                error = ""
            }

            try {
                if (-not $ListOnly) {
                    $one.Publish($pageId, $docxPath, $PublishFormatWord, "")
                    $sectionExported++
                    $totalExported++
                    $entry.status = "ok"
                }
                else {
                    $entry.status = "listed"
                }
            }
            catch {
                $entry.status = "error"
                $entry.error = $_.Exception.Message
                $sectionFailed++
                $totalFailed++
                Write-Warning ("Export fehlgeschlagen fuer '{0}': {1}" -f $title, $_.Exception.Message)

                if ($StopOnError) {
                    $report.Add([pscustomobject]$entry)
                    throw "Abbruch wegen -StopOnError bei Seite '$title' ($pageId)."
                }
            }

            $report.Add([pscustomobject]$entry)
        }

        $totalPages += $pages.Count
        if ($ListOnly) {
            Write-Host ("  Fertig: pages={0}, listed={0}, failed={1}" -f $pages.Count, $sectionFailed)
        }
        else {
            Write-Host ("  Fertig: pages={0}, exported={1}, failed={2}" -f $pages.Count, $sectionExported, $sectionFailed)
        }
    }

    $reportPath = Join-Path $notebookDir "_export-report.json"
    $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $reportPath -Encoding UTF8

    $failures = @($report | Where-Object { $_.status -eq "error" })
    $failureJsonPath = Join-Path $notebookDir "_export-failures.json"
    $failureCsvPath = Join-Path $notebookDir "_export-failures.csv"

    $failures | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $failureJsonPath -Encoding UTF8
    if ($failures.Count -gt 0) {
        $failures |
            Select-Object section, sectionPath, pageIndex, pageId, title, dateTime, output, error |
            Export-Csv -LiteralPath $failureCsvPath -NoTypeInformation -Encoding UTF8
    }
    else {
        Set-Content -LiteralPath $failureCsvPath -Value "section,sectionPath,pageIndex,pageId,title,dateTime,output,error" -Encoding UTF8
    }

    $summaryPath = Join-Path $notebookDir "_export-summary.txt"
    $summaryLines = @(
        "timestamp=$(Get-Date -Format o)",
        "notebook=$($notebookNode.Attributes['name'].Value)",
        "sections=$($requestedSections -join ';')",
        "listOnly=$([bool]$ListOnly)",
        "stopOnError=$([bool]$StopOnError)",
        "totalPages=$totalPages",
        "totalExported=$totalExported",
        "totalFailed=$totalFailed",
        "report=$reportPath",
        "failuresJson=$failureJsonPath",
        "failuresCsv=$failureCsvPath"
    )
    Set-Content -LiteralPath $summaryPath -Value $summaryLines -Encoding UTF8

    Write-Host ""
    Write-Host "Export abgeschlossen."
    Write-Host ("Ausgabe: {0}" -f $notebookDir)
    Write-Host ("Seiten gesamt: {0}" -f $totalPages)
    if ($ListOnly) {
        Write-Host "DOCX gesamt: 0 (ListOnly)"
    }
    else {
        Write-Host ("DOCX gesamt: {0}" -f $totalExported)
    }
    Write-Host ("Fehler gesamt: {0}" -f $totalFailed)
    Write-Host ("Report: {0}" -f $reportPath)
    Write-Host ("Fehler-JSON: {0}" -f $failureJsonPath)
    Write-Host ("Fehler-CSV: {0}" -f $failureCsvPath)
    Write-Host ("Summary: {0}" -f $summaryPath)
    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
