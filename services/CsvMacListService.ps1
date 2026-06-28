<#
.SYNOPSIS
    CSV-Service fuer MAC-Listen.
#>

function Normalize-Mac {
    param([string]$Mac)

    if ([string]::IsNullOrWhiteSpace($Mac)) { return $null }
    return ($Mac -replace '[^0-9A-Fa-f]', '').ToUpper()
}

function Format-MacHyphen {
    param([string]$Mac)

    $norm = Normalize-Mac -Mac $Mac
    if ($null -eq $norm -or $norm.Length -ne 12) { return $Mac }

    return (($norm -split '(.{2})' | Where-Object { $_ -ne '' }) -join '-')
}

function Normalize-MacColumnName {
    param([string]$ColumnName)

    if ([string]::IsNullOrWhiteSpace($ColumnName)) { return $ColumnName }

    if ($ColumnName -ieq 'MacAddress' -or $ColumnName -ieq 'MACAddress') {
        return 'MACAddress'
    }

    return $ColumnName
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
        $columnName = Normalize-MacColumnName -ColumnName ([string]$column)

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
        Normalize-MacColumnName -ColumnName ([string]$_)
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

        $normalizedRequired = Normalize-MacColumnName -ColumnName ([string]$required)
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

function Export-DataTableToMacCsv {
    param(
        [Parameter(Mandatory)] [System.Data.DataTable]$Table,
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [char]$Delimiter
    )

    $export = foreach ($row in $Table.Rows) {
        if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }

        [PSCustomObject]@{
            MACAddress   = [string]$row['MACAddress']
            Description  = [string]$row['Description']
            Description2 = [string]$row['Description2']
            IPAddress    = [string]$row['IPAddress']
            Status       = [string]$row['Status']
        }
    }

    $export | Export-Csv -Path $Path -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
}

function Backup-MacCsv {
    param([Parameter(Mandatory)] [string]$Path)

    if (-not (Test-Path -Path $Path)) { return $null }

    $parent = Split-Path -Path $Path -Parent
    $name = [IO.Path]::GetFileNameWithoutExtension($Path)
    $backupPath = Join-Path $parent ("{0}.bak-{1}.csv" -f $name, (Get-Date -Format 'yyyyMMdd-HHmmss'))

    Copy-Item -Path $Path -Destination $backupPath -Force
    return $backupPath
}

function Test-MacListDataTable {
    param([Parameter(Mandatory)] [System.Data.DataTable]$Table)

    $errors = New-Object System.Collections.Generic.List[string]
    $rowNumber = 1

    foreach ($row in $Table.Rows) {
        if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }

        $mac = [string]$row['MACAddress']
        $ip  = [string]$row['IPAddress']

        $isEmptyRow = [string]::IsNullOrWhiteSpace($mac) -and [string]::IsNullOrWhiteSpace($ip)
        if ($isEmptyRow) {
            $rowNumber++
            continue
        }

        $macNorm = Normalize-Mac -Mac $mac
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
