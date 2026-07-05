<#
.SYNOPSIS
    CSV-Import, Export und Hilfsfunktionen fuer MAC-Listen.

.DESCRIPTION
    Stellt die DataTable-Normalisierung, den CSV-Import/-Export und Backup-Hilfen fuer die DHCP MAC-Reservierung bereit.

.NOTES
    Projekt: DHCP MAC-Reservierung
    Umgebung: Windows Server 2022 / Windows PowerShell 5.1
#>

function ConvertTo-NormalizedMac {
    param([string]$Mac)

    if ([string]::IsNullOrWhiteSpace($Mac)) { return $null }
    return ($Mac -replace '[^0-9A-Fa-f]', '').ToUpper()
}

function Format-MacHyphen {
    param([string]$Mac)

    $norm = ConvertTo-NormalizedMac -Mac $Mac
    if ($null -eq $norm -or $norm.Length -ne 12) { return $Mac }

    return (($norm -split '(.{2})' | Where-Object { $_ -ne '' }) -join '-')
}

function ConvertTo-NormalizedMacColumnName {
    param([string]$ColumnName)

    if ([string]::IsNullOrWhiteSpace($ColumnName)) { return $ColumnName }

    if ($ColumnName -ieq 'MacAddress' -or $ColumnName -ieq 'MACAddress') {
        return 'MACAddress'
    }

    return $ColumnName
}

function New-MacListStatus {
    <#
    .SYNOPSIS
        Erzeugt einen standardisierten Statuswert fuer CSV-Zeilen.

    .DESCRIPTION
        Status wird als Protokollfeld der CSV verwendet. Das Format ist
        LEVEL | yyyy-MM-dd HH:mm:ss | Meldung.

        LEVEL ist einer von OK, FEHLER, WARNUNG oder IMPORT. Es gibt keine
        separaten Logdateien; die Status-Spalte protokolliert den Zustand pro
        Eintrag direkt in der Arbeits-CSV.

    .PARAMETER Level
        Statusstufe (OK, FEHLER, WARNUNG, IMPORT).

    .PARAMETER Message
        Menschlich lesbare Meldung fuer die Zeile.

    .OUTPUTS
        System.String
    #>
    param(
        [ValidateSet('OK','FEHLER','WARNUNG','IMPORT')]
        [string]$Level,

        [Parameter(Mandatory)]
        [string]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    return "$Level | $timestamp | $Message"
}

function Get-CsvDelimiter {
    param(
        [Parameter(Mandatory)] [string]$Path
    )

    if (-not (Test-Path $Path)) {
        throw "CSV-Datei nicht gefunden: $Path"
    }

    $firstLine = Get-Content -Path $Path -TotalCount 1 -ErrorAction Stop

    if ([string]::IsNullOrWhiteSpace($firstLine)) {
        throw "CSV-Datei ist leer oder hat keine Kopfzeile: $Path"
    }

    if ($firstLine -match ';') {
        return ';'
    }

    if ($firstLine -match ',') {
        return ','
    }

    throw "CSV-Trennzeichen konnte nicht erkannt werden. Erste Zeile: $firstLine"
}

function New-MacListDataTable {
    param([string[]]$Columns)

    if ($null -eq $Columns -or $Columns.Count -eq 0) {
        $Columns = @('MACAddress','Description','Description2','IPAddress','Status')
    }

    $table = New-Object System.Data.DataTable

    foreach ($column in $Columns) {
        if ([string]::IsNullOrWhiteSpace([string]$column)) { continue }
        $columnName = ConvertTo-NormalizedMacColumnName -ColumnName ([string]$column)

        if (-not $table.Columns.Contains($columnName)) {
            [void]$table.Columns.Add($columnName)
        }
    }

    if (-not $table.Columns.Contains('Status')) {
        [void]$table.Columns.Add('Status')
    }

    return ,$table
}

function Import-MacCsvToDataTable {
    <#
    .SYNOPSIS
        Liest eine CSV in eine DataTable ein.

    .DESCRIPTION
        Normalisiert die Quellspalten auf die internen Zielspalten und liefert
        Table plus erkanntes Delimiter als PSCustomObject zurück.

        Die DataTable wird absichtlich als unary-comma Rueckgabe verpackt, damit
        PowerShell sie nicht als Zeilenauflistung enumeriert.

    .PARAMETER Path
        Pfad zur CSV-Datei.

    .PARAMETER RequiredColumns
        Erforderliche Quellspalten fuer den Import.

    .OUTPUTS
        PSCustomObject mit den Eigenschaften Table und Delimiter.
    #>
    param(
        [Parameter(Mandatory)] [string]$Path,
        [string[]]$RequiredColumns
    )

    if (-not (Test-Path $Path)) {
        throw "CSV-Datei nicht gefunden: $Path"
    }

    if ($null -eq $RequiredColumns -or $RequiredColumns.Count -eq 0) {
        $RequiredColumns = @(
            'MACAddress',
            'Description',
            'Description2',
            'IPAddress'
        )
    }

    # Zielspalten festlegen, die das Programm intern verwendet.
    $targetColumns = @()
    foreach ($col in $RequiredColumns) {
        if (-not [string]::IsNullOrWhiteSpace($col)) {
            $targetColumns += [string]$col
        }
    }

    if (-not ($targetColumns -contains 'Status')) {
        $targetColumns += 'Status'
    }

    # Wichtig: interne Standardspalte soll MACAddress heißen,
    # auch wenn die CSV MacAddress enthält.
    $targetColumns = $targetColumns | ForEach-Object {
        ConvertTo-NormalizedMacColumnName -ColumnName ([string]$_)
    } | Select-Object -Unique

    $delimiter = Get-CsvDelimiter -Path $Path
    $rows = @(Import-Csv -Path $Path -Delimiter $delimiter)

    $table = New-MacListDataTable -Columns $targetColumns

    if ($rows.Count -eq 0) {
        return [PSCustomObject]@{
            Table     = $table
            Delimiter = $delimiter
        }
    }

    $firstRow = $rows[0]
    if ($null -eq $firstRow) {
        return [PSCustomObject]@{
            Table     = $table
            Delimiter = $delimiter
        }
    }

    $sourceColumns = @()
    if ($firstRow.PSObject -and $firstRow.PSObject.Properties) {
        $sourceColumns = @($firstRow.PSObject.Properties.Name)
    }

    # Alias-Mapping: Quelle -> interne Zielspalte
    $aliases = @{
        'MACAddress' = @('MACAddress', 'MacAddress', 'MAC', 'Mac')
        'Description' = @('Description', 'Name')
        'Description2' = @('Description2', 'Beschreibung', 'Comment')
        'IPAddress' = @('IPAddress', 'IP', 'Adresse')
        'Status' = @('Status')
    }

    function Get-SourceColumnName {
        param(
            [string]$TargetColumn,
            [string[]]$AvailableColumns
        )

        if ($aliases.ContainsKey($TargetColumn)) {
            foreach ($candidate in $aliases[$TargetColumn]) {
                foreach ($available in $AvailableColumns) {
                    if ($available -ieq $candidate) {
                        return $available
                    }
                }
            }
        }

        foreach ($available in $AvailableColumns) {
            if ($available -ieq $TargetColumn) {
                return $available
            }
        }

        return $null
    }

    foreach ($required in $RequiredColumns) {
        if ([string]::IsNullOrWhiteSpace($required)) {
            continue
        }

        $normalizedRequired = ConvertTo-NormalizedMacColumnName -ColumnName ([string]$required)
        $sourceName = Get-SourceColumnName -TargetColumn $normalizedRequired -AvailableColumns $sourceColumns

        if ([string]::IsNullOrWhiteSpace($sourceName)) {
            throw "CSV fehlt erforderliche Spalte: $required. Gefundene Spalten: $($sourceColumns -join ', ')"
        }
    }

    foreach ($row in $rows) {
        if ($null -eq $row) {
            continue
        }

        $dataRow = $table.NewRow()

        foreach ($targetColumn in $targetColumns) {
            $value = ''
            $sourceColumn = Get-SourceColumnName -TargetColumn $targetColumn -AvailableColumns $sourceColumns

            if (-not [string]::IsNullOrWhiteSpace($sourceColumn)) {
                try {
                    $property = $row.PSObject.Properties[$sourceColumn]
                    if ($null -ne $property -and $null -ne $property.Value) {
                        $value = [string]$property.Value
                    }
                } catch {
                    $value = ''
                }
            }

            $dataRow[$targetColumn] = $value
        }

        [void]$table.Rows.Add($dataRow)
    }

    return [PSCustomObject]@{
        Table     = $table
        Delimiter = $delimiter
    }
}

function Resolve-DataTableFromDataSource {
    <#
    .SYNOPSIS
        Liefert aus BindingSource, DataView, DataTable oder Object[] die echte DataTable.

    .DESCRIPTION
        Die Funktion wird benoetigt, weil PowerShell und WinForms DataSource-Werte
        manchmal als BindingSource oder als ein-elementiges Objekt-Array weiterreichen.

        Die Rueckgabe nutzt unary comma, damit eine einzelne DataTable nicht als
        Auflistung von Zeilen aufgeloest wird.

    .PARAMETER DataSource
        Beliebige DataSource aus dem WinForms-Binding.

    .OUTPUTS
        System.Data.DataTable
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$DataSource
    )

    if ($null -eq $DataSource) {
        throw "DataSource ist NULL."
    }

    if ($DataSource -is [System.Data.DataTable]) {
        return ,$DataSource
    }

    if ($DataSource -is [System.Data.DataView]) {
        return ,$DataSource.Table
    }

    if ($DataSource -is [System.Windows.Forms.BindingSource]) {
        return Resolve-DataTableFromDataSource -DataSource $DataSource.DataSource
    }

    if ($DataSource -is [object[]]) {
        foreach ($item in $DataSource) {
            if ($item -is [System.Data.DataTable]) {
                return ,$item
            }

            if ($item -is [System.Data.DataView]) {
                return ,$item.Table
            }
        }

        if ($DataSource.Count -eq 1) {
            return Resolve-DataTableFromDataSource -DataSource $DataSource[0]
        }
    }

    throw "DataSource ist keine DataTable, sondern: $($DataSource.GetType().FullName)"
}

function Import-MacCsvRows {
    param(
        [Parameter(Mandatory)] [string]$Path
    )

    if (-not (Test-Path $Path)) {
        throw "CSV-Datei nicht gefunden: $Path"
    }

    $delimiter = Get-CsvDelimiter -Path $Path
    $rows = @(Import-Csv -Path $Path -Delimiter $delimiter)

    return [PSCustomObject]@{
        Rows      = $rows
        Delimiter = $delimiter
    }
}

function Get-ExistingMacsFromDataTable {
    param(
        [Parameter(Mandatory)] [object]$Table
    )

    $targetTable = @(Resolve-DataTableFromDataSource -DataSource $Table)[0]
    if (-not ($targetTable -is [System.Data.DataTable])) {
        throw "Table ist keine DataTable, sondern: $($targetTable.GetType().FullName)"
    }

    $macSet = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($row in $targetTable.Rows) {
        if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }

        $macNorm = ConvertTo-NormalizedMac -Mac ([string]$row['MACAddress'])
        if (-not [string]::IsNullOrWhiteSpace($macNorm)) {
            [void]$macSet.Add($macNorm)
        }
    }

    return ,$macSet
}

function Get-ExistingIpsFromDataTable {
    param(
        [Parameter(Mandatory)] [object]$Table
    )

    $targetTable = @(Resolve-DataTableFromDataSource -DataSource $Table)[0]
    if (-not ($targetTable -is [System.Data.DataTable])) {
        throw "Table ist keine DataTable, sondern: $($targetTable.GetType().FullName)"
    }

    $usedIpSet = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($row in $targetTable.Rows) {
        if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }

        $ipAddress = $null
        if ([System.Net.IPAddress]::TryParse([string]$row['IPAddress'], [ref]$ipAddress)) {
            [void]$usedIpSet.Add($ipAddress.ToString())
            continue
        }

        $ipText = ([string]$row['IPAddress']).Trim()
        if (-not [string]::IsNullOrWhiteSpace($ipText)) {
            [void]$usedIpSet.Add($ipText)
        }
    }

    return ,$usedIpSet
}

function Get-UsedIpsFromDataTable {
    param(
        [Parameter(Mandatory)] [object]$Table
    )

    return Get-ExistingIpsFromDataTable -Table $Table
}

function Add-MacCsvRowsToDataTable {
    <#
    .SYNOPSIS
        Fuegt CSV-Zeilen als neue Eintraege in eine DataTable ein.

    .DESCRIPTION
        Der Import schreibt nur neue Zeilen in die bestehende Tabelle. DHCP-Aktionen
        werden dabei nicht ausgefuehrt. Der Status der importierten Zeile wird als
        Protokolleintrag im CSV-Statusfeld abgelegt.

    .PARAMETER Table
        Ziel-Datenquelle der Form oder die echte DataTable.

    .PARAMETER ImportPath
        Pfad zur Import-CSV.

    .PARAMETER StartIp
        Untere Grenze fuer die automatische IP-Vergabe.

    .PARAMETER EndIp
        Obere Grenze fuer die automatische IP-Vergabe.

    .PARAMETER Settings
        Projektweite Einstellungen inklusive globalem DHCP-Scope.

    .OUTPUTS
        PSCustomObject mit Summary und der erweiterten DataTable.
    #>
    param(
        [Parameter(Mandatory)] [object]$Table,
        [Parameter(Mandatory)] [string]$ImportPath,
        [Parameter(Mandatory)] [string]$StartIp,
        [Parameter(Mandatory)] [string]$EndIp,
        [Parameter(Mandatory)] [hashtable]$Settings
    )

    $targetTable = @(Resolve-DataTableFromDataSource -DataSource $Table)[0]
    if (-not ($targetTable -is [System.Data.DataTable])) {
        throw "Table ist keine DataTable, sondern: $($targetTable.GetType().FullName)"
    }

    foreach ($columnName in @('MACAddress', 'Description', 'Description2', 'IPAddress')) {
        if (-not $targetTable.Columns.Contains($columnName)) {
            throw "Die Zieltabelle muss die Spalte '$columnName' enthalten."
        }
    }

    if (-not $targetTable.Columns.Contains('Status')) {
        [void]$targetTable.Columns.Add('Status')
    }

    $startIpAddress = $null
    if (-not [System.Net.IPAddress]::TryParse($StartIp, [ref]$startIpAddress) -or
        $startIpAddress.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
        throw 'Start-IP muss eine gueltige IPv4-Adresse sein.'
    }

    $endIpAddress = $null
    if (-not [System.Net.IPAddress]::TryParse($EndIp, [ref]$endIpAddress) -or
        $endIpAddress.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
        throw 'End-IP muss eine gueltige IPv4-Adresse sein.'
    }

    $startValue = Convert-IPv4AddressToUInt32 -IPAddress $startIpAddress
    $endValue = Convert-IPv4AddressToUInt32 -IPAddress $endIpAddress
    if ($startValue -gt $endValue) {
        throw 'Start-IP muss kleiner oder gleich End-IP sein.'
    }

    $existingMacs = Get-ExistingMacsFromDataTable -Table $targetTable
    $usedIps = Get-ExistingIpsFromDataTable -Table $targetTable
    $warnings = New-Object System.Collections.Generic.List[string]

    function Get-ImportValue {
        param(
            [Parameter(Mandatory)] $Row,
            [Parameter(Mandatory)] [string[]]$CandidateNames
        )

        foreach ($candidateName in $CandidateNames) {
            foreach ($property in $Row.PSObject.Properties) {
                if ($property.Name -ieq $candidateName) {
                    if ($null -eq $property.Value) { return '' }
                    return ([string]$property.Value).Trim()
                }
            }
        }

        return ''
    }

    try {
        if ($Settings -and -not [string]::IsNullOrWhiteSpace([string]$Settings.DhcpServer) -and
            -not [string]::IsNullOrWhiteSpace([string]$Settings.DhcpScope)) {
            $reservations = @(Get-DhcpServerv4Reservation -ComputerName $Settings.DhcpServer -ScopeId $Settings.DhcpScope -ErrorAction Stop)
            foreach ($reservation in $reservations) {
                if ($null -ne $reservation -and $reservation.IPAddress) {
                    [void]$usedIps.Add($reservation.IPAddress.ToString())
                }
            }
        }
    }
    catch {
        [void]$warnings.Add('DHCP-Reservierungen konnten nicht geladen werden. Es wurden nur die IPs aus der aktuellen Tabelle beruecksichtigt.')
    }

    $importData = Import-MacCsvRows -Path $ImportPath
    if ($null -eq $importData -or $null -eq $importData.Rows) {
        throw 'Die Importdatei konnte nicht gelesen werden.'
    }

    $imported = 0
    $skippedDuplicateMac = 0
    $skippedInvalidMac = 0
    $skippedNoFreeIp = 0

    foreach ($row in $importData.Rows) {
        if ($null -eq $row) { continue }

        $macRaw = ''
        foreach ($candidateName in @('MACAddress', 'MacAddress')) {
            foreach ($property in $row.PSObject.Properties) {
                if ($property.Name -ieq $candidateName) {
                    if ($null -ne $property.Value) {
                        $macRaw = ([string]$property.Value).Trim()
                    }

                    break
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($macRaw)) { break }
        }

        $macNorm = ConvertTo-NormalizedMac -Mac $macRaw
        if ([string]::IsNullOrWhiteSpace($macNorm) -or $macNorm.Length -ne 12) {
            $skippedInvalidMac++
            continue
        }

        if ($existingMacs.Contains($macNorm)) {
            $skippedDuplicateMac++
            continue
        }

        $freeIp = Get-NextFreeIpInRange -StartIp $StartIp -EndIp $EndIp -UsedIps $usedIps
        if ([string]::IsNullOrWhiteSpace($freeIp)) {
            $skippedNoFreeIp++
            continue
        }

        $newRow = $targetTable.NewRow()
        $newRow['MACAddress'] = Format-MacHyphen -Mac $macNorm
        $newRow['Description'] = Get-ImportValue -Row $row -CandidateNames @('Description')
        $newRow['Description2'] = Get-ImportValue -Row $row -CandidateNames @('Description2')
        $newRow['IPAddress'] = $freeIp
        $newRow['Status'] = New-MacListStatus -Level 'IMPORT' -Message 'Importiert - noch nicht reserviert'
        [void]$targetTable.Rows.Add($newRow)

        [void]$existingMacs.Add($macNorm)
        [void]$usedIps.Add($freeIp)
        $imported++
    }

    return [PSCustomObject]@{
        Table               = $targetTable
        ImportPath          = $ImportPath
        Imported            = $imported
        SkippedDuplicateMac = $skippedDuplicateMac
        SkippedInvalidMac   = $skippedInvalidMac
        SkippedNoFreeIp     = $skippedNoFreeIp
        Warnings            = @($warnings)
    }
}

function Export-DataTableToMacCsv {
    <#
    .SYNOPSIS
        Exportiert eine DataTable in die MAC-CSV.

    .DESCRIPTION
        Es werden nur die Spalten MACAddress, Description, Description2, IPAddress
        und Status geschrieben. Gelöschte und komplett leere Zeilen werden uebersprungen.

        Die Status-Spalte bleibt als Protokoll pro Eintrag erhalten.

    .PARAMETER Table
        Die zu exportierende DataTable oder eine normalisierbare DataSource.

    .PARAMETER Path
        Zieldatei fuer den CSV-Export.

    .PARAMETER Delimiter
        Verwendetes CSV-Trennzeichen.
    #>
    param(
        [Parameter(Mandatory)] [object]$Table,
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [char]$Delimiter
    )

    $targetTable = @(Resolve-DataTableFromDataSource -DataSource $Table)[0]
    if (-not ($targetTable -is [System.Data.DataTable])) {
        throw "Table ist keine DataTable, sondern: $($targetTable.GetType().FullName)"
    }

    $export = foreach ($row in $targetTable.Rows) {
        if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }

        $mac = [string]$row['MACAddress']
        $name = [string]$row['Description']
        $description = [string]$row['Description2']
        $ip = [string]$row['IPAddress']
        $status = [string]$row['Status']

        if ([string]::IsNullOrWhiteSpace($mac) -and
            [string]::IsNullOrWhiteSpace($name) -and
            [string]::IsNullOrWhiteSpace($description) -and
            [string]::IsNullOrWhiteSpace($ip) -and
            [string]::IsNullOrWhiteSpace($status)) {
            continue
        }

        [PSCustomObject]@{
            MACAddress   = $mac
            Description  = $name
            Description2 = $description
            IPAddress    = $ip
            Status       = $status
        }
    }

    $export | Export-Csv -Path $Path -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
}

function Backup-MacCsv {
    <#
    .SYNOPSIS
        Erstellt ein Backup der aktuellen CSV im Unterordner Backups.

    .DESCRIPTION
        Backups werden unter <MacListenPfad>\Backups mit Originalnamen und Timestamp
        gespeichert. Es werden keine zusätzlichen Logdateien erzeugt; der Verlauf
        steht pro Eintrag in der Status-Spalte.
    #>
    param([Parameter(Mandatory)] [string]$Path)

    if (-not (Test-Path -Path $Path)) { return $null }

    $parent = Split-Path -Path $Path -Parent
    $backupFolder = Join-Path $parent 'Backups'
    [void](New-Item -ItemType Directory -Path $backupFolder -Force)
    $name = [IO.Path]::GetFileNameWithoutExtension($Path)
    $backupPath = Join-Path $backupFolder ("{0}_{1}.csv" -f $name, (Get-Date -Format 'yyyyMMdd-HHmmss'))

    Copy-Item -Path $Path -Destination $backupPath -Force
    return $backupPath
}

function Test-MacListDataTable {
    param([Parameter(Mandatory)] [object]$Table)

    $targetTable = @(Resolve-DataTableFromDataSource -DataSource $Table)[0]
    if (-not ($targetTable -is [System.Data.DataTable])) {
        throw "Table ist keine DataTable, sondern: $($targetTable.GetType().FullName)"
    }

    $errors = New-Object System.Collections.Generic.List[string]
    $rowNumber = 1

    foreach ($row in $targetTable.Rows) {
        if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }

        $mac = [string]$row['MACAddress']
        $ip  = [string]$row['IPAddress']

        $isEmptyRow = [string]::IsNullOrWhiteSpace($mac) -and [string]::IsNullOrWhiteSpace($ip)
        if ($isEmptyRow) {
            $rowNumber++
            continue
        }

        $macNorm = ConvertTo-NormalizedMac -Mac $mac
        if (-not $macNorm -or $macNorm.Length -ne 12) {
            $errors.Add("Zeile ${rowNumber}: Ungueltige MAC-Adresse '$mac'")
        }

        $ipObj = $null
        if (-not [System.Net.IPAddress]::TryParse($ip, [ref]$ipObj)) {
            $errors.Add("Zeile ${rowNumber}: Ungueltige IP-Adresse '$ip'")
        }

        $rowNumber++
    }

    return $errors.ToArray()
}
