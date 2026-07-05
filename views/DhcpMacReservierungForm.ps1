<#
.SYNOPSIS
    WinForms-Oberflaeche fuer die DHCP MAC-Reservierung.

.DESCRIPTION
    Stellt die CSV-basierte Arbeitsoberflaeche bereit, in der MAC-, Name-,
    Beschreibung- und IP-Daten bearbeitet, validiert und gespeichert werden.

.NOTES
    Projekt: DHCP MAC-Reservierung
    Umgebung: Windows Server 2022 / Windows PowerShell 5.1
#>

<#
.SYNOPSIS
    Zeigt die WinForms-Oberflaeche an.

.PARAMETER Settings
    Projektweite Einstellungen aus Settings.psd1.

.OUTPUTS
    Keine.
#>
function Show-DhcpMacReservierungForm {
    param([Parameter(Mandatory)] [hashtable]$Settings)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'DHCP MAC-Reservierung'
    $form.ClientSize = New-Object System.Drawing.Size(1000,720)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(900,600)
    $form.Padding = New-Object System.Windows.Forms.Padding(12)

    $fontDefault = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.Font = $fontDefault

    $groupSource = New-Object System.Windows.Forms.GroupBox
    $groupSource.Location = New-Object System.Drawing.Point(12,12)
    $groupSource.Size = New-Object System.Drawing.Size(976,95)
    $groupSource.Anchor = 'Top, Left, Right'
    $groupSource.Text = 'Quelle'

    $form.Controls.Add($groupSource)

    $lblArea = New-Object System.Windows.Forms.Label
    $lblArea.Location = New-Object System.Drawing.Point(16,28)
    $lblArea.Size = New-Object System.Drawing.Size(120,22)
    $lblArea.Text = 'Bereichsvorlage:'
    $groupSource.Controls.Add($lblArea)

    $comboArea = New-Object System.Windows.Forms.ComboBox
    $comboArea.Location = New-Object System.Drawing.Point(145,24)
    $comboArea.Size = New-Object System.Drawing.Size(360,24)
    $comboArea.Anchor = 'Top, Left, Right'
    $comboArea.DropDownStyle = 'DropDownList'
    $groupSource.Controls.Add($comboArea)

    $lblCsvFolder = New-Object System.Windows.Forms.Label
    $lblCsvFolder.Location = New-Object System.Drawing.Point(16,54)
    $lblCsvFolder.Size = New-Object System.Drawing.Size(120,22)
    $lblCsvFolder.Text = 'CSV:'
    $groupSource.Controls.Add($lblCsvFolder)

    $txtCsvFolder = New-Object System.Windows.Forms.TextBox
    $txtCsvFolder.Location = New-Object System.Drawing.Point(145,52)
    $txtCsvFolder.Size = New-Object System.Drawing.Size(470,24)
    $txtCsvFolder.Anchor = 'Top, Left, Right'
    $groupSource.Controls.Add($txtCsvFolder)

    $comboCsvFile = New-Object System.Windows.Forms.ComboBox
    $comboCsvFile.Location = New-Object System.Drawing.Point(620,52)
    $comboCsvFile.Size = New-Object System.Drawing.Size(240,24)
    $comboCsvFile.Anchor = 'Top, Left, Right'
    $comboCsvFile.DropDownStyle = 'DropDownList'
    $groupSource.Controls.Add($comboCsvFile)

    $btnBrowseFolder = New-Object System.Windows.Forms.Button
    $btnBrowseFolder.Location = New-Object System.Drawing.Point(868,50)
    $btnBrowseFolder.Size = New-Object System.Drawing.Size(40,28)
    $btnBrowseFolder.Anchor = 'Top, Right'
    $btnBrowseFolder.Text = '...'
    $groupSource.Controls.Add($btnBrowseFolder)

    $groupIp = New-Object System.Windows.Forms.GroupBox
    $groupIp.Location = New-Object System.Drawing.Point(12,117)
    $groupIp.Size = New-Object System.Drawing.Size(976,70)
    $groupIp.Anchor = 'Top, Left, Right'
    $groupIp.Text = 'IP-Bereich'
    $form.Controls.Add($groupIp)

    $lblScope = New-Object System.Windows.Forms.Label
    $lblScope.Location = New-Object System.Drawing.Point(16,29)
    $lblScope.Size = New-Object System.Drawing.Size(60,22)
    $lblScope.Text = 'Scope:'
    $groupIp.Controls.Add($lblScope)

    $txtScope = New-Object System.Windows.Forms.TextBox
    $txtScope.Location = New-Object System.Drawing.Point(82,25)
    $txtScope.Size = New-Object System.Drawing.Size(150,24)
    $txtScope.ReadOnly = $true
    $txtScope.TabStop = $false
    $groupIp.Controls.Add($txtScope)

    $lblStartIp = New-Object System.Windows.Forms.Label
    $lblStartIp.Location = New-Object System.Drawing.Point(258,29)
    $lblStartIp.Size = New-Object System.Drawing.Size(60,22)
    $lblStartIp.Text = 'Start-IP:'
    $groupIp.Controls.Add($lblStartIp)

    $txtStartIp = New-Object System.Windows.Forms.TextBox
    $txtStartIp.Location = New-Object System.Drawing.Point(324,25)
    $txtStartIp.Size = New-Object System.Drawing.Size(150,24)
    $groupIp.Controls.Add($txtStartIp)

    $lblEndIp = New-Object System.Windows.Forms.Label
    $lblEndIp.Location = New-Object System.Drawing.Point(492,29)
    $lblEndIp.Size = New-Object System.Drawing.Size(55,22)
    $lblEndIp.Text = 'End-IP:'
    $groupIp.Controls.Add($lblEndIp)

    $txtEndIp = New-Object System.Windows.Forms.TextBox
    $txtEndIp.Location = New-Object System.Drawing.Point(552,25)
    $txtEndIp.Size = New-Object System.Drawing.Size(150,24)
    $groupIp.Controls.Add($txtEndIp)

    $status = New-Object System.Windows.Forms.Label
    $status.Location = New-Object System.Drawing.Point(12,627)
    $status.Size = New-Object System.Drawing.Size(976,22)
    $status.Anchor = 'Bottom, Left, Right'
    $status.Text = 'Bereit.'
    $form.Controls.Add($status)

    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Location = New-Object System.Drawing.Point(12,652)
    $bottomPanel.Size = New-Object System.Drawing.Size(976,58)
    $bottomPanel.Anchor = 'Bottom, Left, Right'
    $form.Controls.Add($bottomPanel)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12,195)
    $grid.Size = New-Object System.Drawing.Size(976,425)
    $grid.Anchor = 'Top, Bottom, Left, Right'
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.AllowUserToAddRows = $true
    $grid.AllowUserToDeleteRows = $true
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
    $grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableWithoutHeaderText
    $grid.MultiSelect = $false
    $grid.RowHeadersVisible = $true
    $grid.RowHeadersWidth = 55
    $grid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
    $grid.EditMode = 'EditOnKeystrokeOrF2'
    $grid.AutoGenerateColumns = $true
    $form.Controls.Add($grid)

    $bindingSource = New-Object System.Windows.Forms.BindingSource
    $grid.DataSource = $bindingSource

    $grid.Add_KeyDown({
        param($eventSender, $e)
        [void]$eventSender

        if (-not $e.Control -and -not $e.Shift) { return }

        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
            $targetCell = $grid.CurrentCell
            if ($null -eq $targetCell) {
                return
            }

            $rowIndex = $targetCell.RowIndex
            if ($rowIndex -lt 0 -or $rowIndex -ge $grid.Rows.Count) {
                return
            }

            $targetRow = $grid.Rows[$rowIndex]
            if ($null -eq $targetRow -or $targetRow.IsNewRow) {
                return
            }

            $confirm = [System.Windows.Forms.MessageBox]::Show('Aktuelle Zeile wirklich löschen?', 'Zeile löschen', 'YesNo', 'Question')
            if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
                $grid.Rows.RemoveAt($rowIndex)
                Update-DataGridRowNumbers -Grid $grid
                $grid.Invalidate()
            }

            $e.Handled = $true
            $e.SuppressKeyPress = $true
            return
        }

        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::C) {
            if ($grid.SelectedCells.Count -eq 1) {
                $cell = $grid.SelectedCells[0]
                $value = ''
                if ($null -ne $cell.Value) {
                    $value = [string]$cell.Value
                }

                [System.Windows.Forms.Clipboard]::SetText($value)
                $e.Handled = $true
                $e.SuppressKeyPress = $true
                return
            }

            return
        }

        if ($e.KeyCode -ne [System.Windows.Forms.Keys]::V) {
            return
        }

        $targetCell = $grid.CurrentCell
        if ($null -eq $targetCell -or $targetCell.ReadOnly) {
            return
        }

        $text = [System.Windows.Forms.Clipboard]::GetText()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return
        }

        $parts = $text -split "[`r`n`t]+"
        $value = if ($parts.Count -gt 0) { $parts[0] } else { $text }

        $targetCell.Value = $value
        $grid.NotifyCurrentCellDirty($true)

        $e.Handled = $true
        $e.SuppressKeyPress = $true
    })

    $grid.Add_CellEndEdit({
        Update-ValidationDisplay | Out-Null
    })

    $grid.Add_CellEnter({
        param($eventSender, $e)
        [void]$eventSender

        try {
            Set-NextFreeIpForCell -Grid $grid -RowIndex $e.RowIndex -ColumnIndex $e.ColumnIndex
        }
        catch {
            Set-StatusText "Keine IP vorgeschlagen: $($_.Exception.Message)"
        }
    })

    $grid.Add_DataBindingComplete({
        Update-DataGridRowNumbers -Grid $grid
    })

    $grid.Add_RowsAdded({
        Update-DataGridRowNumbers -Grid $grid
    })

    $grid.Add_RowsRemoved({
        Update-DataGridRowNumbers -Grid $grid
    })

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Location = New-Object System.Drawing.Point(0,14)
    $btnLoad.Size = New-Object System.Drawing.Size(110,30)
    $btnLoad.Anchor = 'Top, Left'
    $btnLoad.Text = 'CSV importieren...'
    $bottomPanel.Controls.Add($btnLoad)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(120,14)
    $btnSave.Size = New-Object System.Drawing.Size(120,30)
    $btnSave.Anchor = 'Top, Left'
    $btnSave.Text = 'CSV speichern'
    $bottomPanel.Controls.Add($btnSave)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(700,14)
    $btnRun.Size = New-Object System.Drawing.Size(170,30)
    $btnRun.Anchor = 'Top, Right'
    $btnRun.Text = 'Reservierungen ausführen'
    $bottomPanel.Controls.Add($btnRun)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(882,14)
    $btnClose.Size = New-Object System.Drawing.Size(82,30)
    $btnClose.Anchor = 'Top, Right'
    $btnClose.Text = 'Schließen'
    $bottomPanel.Controls.Add($btnClose)

    $txtCsvFolder.Text = [string]$Settings.MacListenPfad
    $txtScope.Text = [string]$Settings.DhcpScope

    $script:CurrentCsvPath = $null
    $script:CurrentDelimiter = ';'
    $script:CurrentAreaConfig = $null
    $script:CachedDhcpReservedIpsKey = $null
    $script:CachedDhcpReservedIps = $null
    $script:CachedDhcpReservedIpsWarning = $null

    function Set-StatusText {
        param([string]$Text)
        $status.Text = $Text
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Get-OrderedAreaNames {
        $orderedAreas = foreach ($entry in $Settings.Areas.GetEnumerator()) {
            [pscustomobject]@{
                Name  = [string]$entry.Key
                Order = if ($entry.Value -and $entry.Value.ContainsKey('Order')) { [int]$entry.Value.Order } else { [int]::MaxValue }
            }
        }

        return $orderedAreas | Sort-Object Order, Name | Select-Object -ExpandProperty Name
    }

    function Get-SelectedAreaConfig {
        if (-not $comboArea.SelectedItem) { return $null }
        return $Settings.Areas[[string]$comboArea.SelectedItem]
    }

    function Get-CsvFolderPath {
        return [string]$txtCsvFolder.Text.Trim()
    }

    function Get-CurrentCsvFileName {
        if ($comboCsvFile.SelectedItem) {
            return [string]$comboCsvFile.SelectedItem
        }

        $area = Get-SelectedAreaConfig
        if ($area -and $area.ContainsKey('Csv')) {
            return [string]$area.Csv
        }

        return $null
    }

    function Get-SelectedAreaCsvFileName {
        $area = Get-SelectedAreaConfig
        if ($area -and $area.ContainsKey('Csv')) {
            return [string]$area.Csv
        }

        return $null
    }

    function Get-CurrentCsvPath {
        $folderPath = Get-CsvFolderPath
        $csvFileName = Get-CurrentCsvFileName

        if ([string]::IsNullOrWhiteSpace($folderPath) -or [string]::IsNullOrWhiteSpace($csvFileName)) {
            return $null
        }

        return Join-Path -Path $folderPath -ChildPath $csvFileName
    }

    # Der DHCP-Scope ist global in Settings.DhcpScope definiert.
    # Die Areas liefern nur organisatorische Start-/Endbereiche fuer Vorschlaege und Validierung.

    function Get-DhcpReservedIpsForCurrentScope {
        $serverName = [string]$Settings.DhcpServer
        $scopeText = [string]$Settings.DhcpScope
        $cacheKey = "$serverName|$scopeText"

        if ($script:CachedDhcpReservedIpsKey -eq $cacheKey -and $null -ne $script:CachedDhcpReservedIps) {
            return [PSCustomObject]@{
                Ips     = $script:CachedDhcpReservedIps
                Warning = $script:CachedDhcpReservedIpsWarning
            }
        }

        $reservedIps = New-Object 'System.Collections.Generic.HashSet[string]'
        $warning = $null

        try {
            if (-not [string]::IsNullOrWhiteSpace($serverName) -and -not [string]::IsNullOrWhiteSpace($scopeText)) {
                $reservations = @(Get-DhcpServerv4Reservation -ComputerName $serverName -ScopeId $scopeText -ErrorAction Stop)

                foreach ($reservation in $reservations) {
                    if ($null -ne $reservation -and $reservation.IPAddress) {
                        [void]$reservedIps.Add([string]$reservation.IPAddress.ToString())
                    }
                }
            }
        }
        catch {
            $warning = 'DHCP-Reservierungen konnten nicht geprüft werden; Vorschlag basiert nur auf aktueller Tabelle.'
        }

        $script:CachedDhcpReservedIpsKey = $cacheKey
        $script:CachedDhcpReservedIps = $reservedIps
        $script:CachedDhcpReservedIpsWarning = $warning

        return [PSCustomObject]@{
            Ips     = $reservedIps
            Warning = $warning
        }
    }

    function Get-UsedIpsForSuggestion {
        param(
            [Parameter(Mandatory)] [System.Data.DataTable]$Table,
            [AllowEmptyString()]
            [string]$CurrentIpValue = ''
        )

        $usedIps = Get-ExistingIpsFromDataTable -Table $Table

        if (-not [string]::IsNullOrWhiteSpace($CurrentIpValue)) {
            try {
                [void]$usedIps.Remove([string]$CurrentIpValue.Trim())
            }
            catch {
                [void]$_.Exception
                # Ignorieren, wenn die aktuelle IP nicht vorhanden ist.
            }
        }

        $reservedIpsResult = Get-DhcpReservedIpsForCurrentScope
        if ($null -ne $reservedIpsResult -and $null -ne $reservedIpsResult.Ips) {
            foreach ($ip in $reservedIpsResult.Ips) {
                if (-not [string]::IsNullOrWhiteSpace([string]$ip)) {
                    [void]$usedIps.Add([string]$ip)
                }
            }
        }

        return [PSCustomObject]@{
            Ips     = $usedIps
            Warning = if ($null -ne $reservedIpsResult) { $reservedIpsResult.Warning } else { $null }
        }
    }

    function Set-NextFreeIpForCell {
        param(
            [Parameter(Mandatory)] [System.Windows.Forms.DataGridView]$Grid,
            [Parameter(Mandatory)] [int]$RowIndex,
            [Parameter(Mandatory)] [int]$ColumnIndex
        )

        if ($RowIndex -lt 0 -or $ColumnIndex -lt 0) { return }
        if ($ColumnIndex -ge $Grid.Columns.Count) { return }

        $column = $Grid.Columns[$ColumnIndex]
        if ($null -eq $column -or $column.Name -ne 'IPAddress') { return }

        $row = $Grid.Rows[$RowIndex]
        if ($null -eq $row) { return }

        $cell = $row.Cells[$ColumnIndex]
        if ($null -eq $cell -or $cell.ReadOnly) { return }

        $currentText = [string]$cell.Value
        if (-not [string]::IsNullOrWhiteSpace($currentText)) { return }

        $startText = ([string]$txtStartIp.Text).Trim()
        $endText = ([string]$txtEndIp.Text).Trim()

        $startIp = Convert-TextToIPv4Address -Text $startText
        $endIp = Convert-TextToIPv4Address -Text $endText
        if ($null -eq $startIp -or $null -eq $endIp) {
            Set-StatusText 'Keine IP vorgeschlagen: Start-IP oder End-IP ungültig.'
            return
        }

        $table = @(Resolve-DataTableFromDataSource -DataSource $bindingSource.DataSource)[0]
        if (-not ($table -is [System.Data.DataTable])) {
            Set-StatusText 'Keine IP vorgeschlagen: aktuelle Tabelle ist keine DataTable.'
            return
        }

        $usedIpsResult = Get-UsedIpsForSuggestion -Table $table -CurrentIpValue $currentText
        $nextIp = Get-NextFreeIpInRange -StartIp $startText -EndIp $endText -UsedIps $usedIpsResult.Ips

        if ([string]::IsNullOrWhiteSpace($nextIp)) {
            Set-StatusText 'Keine freie IP gefunden.'
            return
        }

        try {
            $cell.Value = $nextIp
            $bindingSource.EndEdit()
            $Grid.NotifyCurrentCellDirty($true)
            Update-ValidationDisplay | Out-Null
            Update-DataGridRowNumbers -Grid $Grid

            if (-not [string]::IsNullOrWhiteSpace([string]$usedIpsResult.Warning)) {
                Set-StatusText $usedIpsResult.Warning
            } else {
                Set-StatusText "Nächste freie IP vorgeschlagen: $nextIp"
            }
        }
        catch {
            Set-StatusText "Keine IP vorgeschlagen: $($_.Exception.Message)"
        }
    }

    function Convert-TextToIPv4Address {
        param([string]$Text)

        if ([string]::IsNullOrWhiteSpace($Text)) { return $null }

        $ipAddress = $null
        if (-not [System.Net.IPAddress]::TryParse($Text.Trim(), [ref]$ipAddress)) {
            return $null
        }

        if ($ipAddress.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
            return $null
        }

        return $ipAddress
    }

    function Convert-IPv4ToUInt32 {
        param([Parameter(Mandatory)] [System.Net.IPAddress]$IPAddress)

        $bytes = $IPAddress.GetAddressBytes()
        if ($bytes.Length -ne 4) {
            throw 'Nur IPv4-Adressen werden unterstuetzt.'
        }

        [Array]::Reverse($bytes)
        return [BitConverter]::ToUInt32($bytes, 0)
    }

    function Add-ValidationMessage {
        param(
            [Parameter(Mandatory)] [hashtable]$MessageMap,
            [Parameter(Mandatory)] [int]$RowNumber,
            [Parameter(Mandatory)] [string]$Message
        )

        if (-not $MessageMap.ContainsKey($RowNumber)) {
            $MessageMap[$RowNumber] = New-Object System.Collections.Generic.List[string]
        }

        [void]$MessageMap[$RowNumber].Add($Message)
    }

    function Get-FriendlyFieldName {
        param([Parameter(Mandatory)] [string]$FieldName)

        switch ($FieldName) {
            'MACAddress' { return 'MAC-Adresse' }
            'Description' { return 'Name' }
            'Description2' { return 'Beschreibung' }
            'IPAddress' { return 'IP-Adresse' }
            default { return $FieldName }
        }
    }

    function Get-ValidationContextText {
        param([Parameter(Mandatory)] $Row)

        $name = ([string]$Row['Description']).Trim()
        $ip = ([string]$Row['IPAddress']).Trim()

        if (-not [string]::IsNullOrWhiteSpace($name) -and -not [string]::IsNullOrWhiteSpace($ip)) {
            return "($name, $ip)"
        }

        if (-not [string]::IsNullOrWhiteSpace($name)) {
            return "($name)"
        }

        if (-not [string]::IsNullOrWhiteSpace($ip)) {
            return "($ip)"
        }

        return $null
    }

    function Add-ValidationFieldMessage {
        param(
            [Parameter(Mandatory)] [PSCustomObject]$Report,
            [Parameter(Mandatory)] [hashtable]$MessageMap,
            [Parameter(Mandatory)] [int]$RowNumber,
            [Parameter(Mandatory)] [string]$FieldName,
            [Parameter(Mandatory)] [string]$Message,
            [Parameter(Mandatory)] $Row
        )

        $contextText = Get-ValidationContextText -Row $Row
        $prefix = "Zeile ${RowNumber}"
        if ($contextText) {
            $prefix += " $contextText"
        }

        $friendlyFieldName = Get-FriendlyFieldName -FieldName $FieldName
        $formattedMessage = "$($prefix): $friendlyFieldName $Message"
        [void]$Report.Errors.Add($formattedMessage)
        Add-ValidationMessage -MessageMap $MessageMap -RowNumber $RowNumber -Message $formattedMessage
    }

    function Add-ValidationWarningMessage {
        param(
            [Parameter(Mandatory)] [PSCustomObject]$Report,
            [Parameter(Mandatory)] [hashtable]$MessageMap,
            [Parameter(Mandatory)] [int]$RowNumber,
            [Parameter(Mandatory)] [string]$FieldName,
            [Parameter(Mandatory)] [string]$Message,
            [Parameter(Mandatory)] $Row
        )

        $contextText = Get-ValidationContextText -Row $Row
        $prefix = "Zeile ${RowNumber}"
        if ($contextText) {
            $prefix += " $contextText"
        }

        $friendlyFieldName = Get-FriendlyFieldName -FieldName $FieldName
        $formattedMessage = "$($prefix): $friendlyFieldName $Message"
        [void]$Report.Warnings.Add($formattedMessage)
        Add-ValidationMessage -MessageMap $MessageMap -RowNumber $RowNumber -Message $formattedMessage
    }

    function Format-ValidationMessageBoxText {
        param(
            [Parameter(Mandatory)] [string]$Headline,
            [Parameter(Mandatory)] [string[]]$Messages,
            [int]$MaxItems = 20,
            [Parameter(Mandatory)] [string]$ItemLabel
        )

        $lines = New-Object System.Collections.Generic.List[string]
        [void]$lines.Add($Headline)
        [void]$lines.Add('')

        $limit = [Math]::Min($Messages.Count, $MaxItems)
        for ($index = 0; $index -lt $limit; $index++) {
            [void]$lines.Add($Messages[$index])
        }

        if ($Messages.Count -gt $MaxItems) {
            [void]$lines.Add('')
            [void]$lines.Add("... weitere $($Messages.Count - $MaxItems) $ItemLabel nicht angezeigt.")
        }

        return ($lines -join [Environment]::NewLine)
    }

    function Get-DescriptionValidationMessage {
        param([string]$Description)

        $text = [string]$Description
        if ([string]::IsNullOrWhiteSpace($text)) {
            return 'fehlt.'
        }

        $trimmedText = $text.Trim()
        if ($text -ne $trimmedText) {
            return 'enthält führende oder nachgestellte Leerzeichen.'
        }

        if ($text -match ' ') {
            return 'darf keine Leerzeichen enthalten.'
        }

        if ($text.Length -gt 255) {
            return 'ist länger als 255 Zeichen.'
        }

        if ($text -match '[\t\r\n]' -or $text -match '[\x00-\x1F\x7F]') {
            return 'enthält ungültige Zeichen.'
        }

        if ($text -notmatch '^[A-Za-z0-9_\-\(\)]+$') {
            return 'enthält ungültige Zeichen.'
        }

        return $null
    }

    function Get-ValidationReport {
        param(
            [Parameter(Mandatory)] [object]$Table,
            [Parameter(Mandatory)] [string]$StartIpText,
            [Parameter(Mandatory)] [string]$EndIpText
        )

        $targetTable = @(Resolve-DataTableFromDataSource -DataSource $Table)[0]

        $report = [PSCustomObject]@{
            Errors      = New-Object System.Collections.Generic.List[string]
            Warnings    = New-Object System.Collections.Generic.List[string]
            ErrorRows   = @{}
            WarningRows = @{}
            TargetTable = $targetTable
        }

        if (-not ($targetTable -is [System.Data.DataTable])) {
            $typeName = if ($null -eq $targetTable) { 'null' } else { $targetTable.GetType().FullName }
            [void]$report.Errors.Add("Tabelle ist keine DataTable, sondern: $typeName")
            return $report
        }

        if ($null -eq $targetTable) {
            [void]$report.Errors.Add('Keine CSV-Daten geladen.')
            return $report
        }

        $scopeText = ([string]$Settings.DhcpScope).Trim()
        $scopeIp = $null
        if ([string]::IsNullOrWhiteSpace($scopeText) -or -not [System.Net.IPAddress]::TryParse($scopeText, [ref]$scopeIp)) {
            [void]$report.Errors.Add('DHCP-Scope muss eine gueltige IPv4-Adresse sein.')
        } elseif ($scopeIp.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
            [void]$report.Errors.Add('DHCP-Scope muss eine gueltige IPv4-Adresse sein.')
        }

        $startIp = Convert-TextToIPv4Address -Text $StartIpText
        if ($null -eq $startIp) {
            [void]$report.Errors.Add('Start-IP muss eine gueltige IPv4-Adresse sein.')
        }

        $endIp = Convert-TextToIPv4Address -Text $EndIpText
        if ($null -eq $endIp) {
            [void]$report.Errors.Add('End-IP muss eine gueltige IPv4-Adresse sein.')
        }

        $startValue = $null
        $endValue = $null
        if ($report.Errors.Count -eq 0) {
            $startValue = Convert-IPv4ToUInt32 -IPAddress $startIp
            $endValue = Convert-IPv4ToUInt32 -IPAddress $endIp

            if ($startValue -gt $endValue) {
                [void]$report.Errors.Add('Start-IP muss kleiner oder gleich End-IP sein.')
            }
        }

        $macCounts = @{}
        $macRows = @{}
        $ipCounts = @{}
        $ipRows = @{}
        $descriptionCounts = @{}
        $descriptionRows = @{}

        $rowNumber = 1
        foreach ($row in $targetTable.Rows) {
            if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }

            $macRaw = [string]$row['MACAddress']
            $ipRaw = [string]$row['IPAddress']
            $descriptionRaw = [string]$row['Description']
            $description2Raw = [string]$row['Description2']

            $isEmptyRow = [string]::IsNullOrWhiteSpace($macRaw) -and [string]::IsNullOrWhiteSpace($ipRaw) -and [string]::IsNullOrWhiteSpace($descriptionRaw) -and [string]::IsNullOrWhiteSpace($description2Raw)
            if ($isEmptyRow) {
                $rowNumber++
                continue
            }

            $descriptionMessage = Get-DescriptionValidationMessage -Description $descriptionRaw
            if ($descriptionMessage) {
                Add-ValidationFieldMessage -Report $report -MessageMap $report.ErrorRows -RowNumber $rowNumber -FieldName 'Description' -Message $descriptionMessage -Row $row
            }

            $macNorm = ConvertTo-NormalizedMac -Mac $macRaw
            if (-not [string]::IsNullOrWhiteSpace($macRaw) -and (-not $macNorm -or $macNorm.Length -ne 12)) {
                Add-ValidationFieldMessage -Report $report -MessageMap $report.ErrorRows -RowNumber $rowNumber -FieldName 'MACAddress' -Message "ist ungültig: $macRaw" -Row $row
            } elseif (-not [string]::IsNullOrWhiteSpace($macNorm)) {
                if (-not $macCounts.ContainsKey($macNorm)) {
                    $macCounts[$macNorm] = 0
                    $macRows[$macNorm] = New-Object System.Collections.Generic.List[int]
                }

                $macCounts[$macNorm]++
                [void]$macRows[$macNorm].Add($rowNumber)
            }

            $ipAddress = $null
            if (-not [string]::IsNullOrWhiteSpace($ipRaw)) {
                if (-not [System.Net.IPAddress]::TryParse($ipRaw, [ref]$ipAddress)) {
                    Add-ValidationFieldMessage -Report $report -MessageMap $report.ErrorRows -RowNumber $rowNumber -FieldName 'IPAddress' -Message "ist ungültig: $ipRaw" -Row $row
                } else {
                    $ipText = $ipAddress.ToString()
                    if (-not $ipCounts.ContainsKey($ipText)) {
                        $ipCounts[$ipText] = 0
                        $ipRows[$ipText] = New-Object System.Collections.Generic.List[int]
                    }

                    $ipCounts[$ipText]++
                    [void]$ipRows[$ipText].Add($rowNumber)

                    if ($null -ne $startValue -and $null -ne $endValue) {
                        $ipValue = Convert-IPv4ToUInt32 -IPAddress $ipAddress
                        if ($ipValue -lt $startValue -or $ipValue -gt $endValue) {
                            Add-ValidationFieldMessage -Report $report -MessageMap $report.ErrorRows -RowNumber $rowNumber -FieldName 'IPAddress' -Message "liegt außerhalb des Bereichs $StartIpText - $EndIpText" -Row $row
                        }
                    }
                }
            }

            $descriptionNorm = if ([string]::IsNullOrWhiteSpace($descriptionRaw)) { $null } else { $descriptionRaw.Trim().ToLowerInvariant() }
            if (-not [string]::IsNullOrWhiteSpace($descriptionNorm)) {
                if (-not $descriptionCounts.ContainsKey($descriptionNorm)) {
                    $descriptionCounts[$descriptionNorm] = 0
                    $descriptionRows[$descriptionNorm] = New-Object System.Collections.Generic.List[int]
                }

                $descriptionCounts[$descriptionNorm]++
                [void]$descriptionRows[$descriptionNorm].Add($rowNumber)
            }

            $rowNumber++
        }

        foreach ($entry in $descriptionCounts.GetEnumerator()) {
            if ($entry.Value -le 1) { continue }

            foreach ($duplicateRow in $descriptionRows[$entry.Key]) {
                $row = $targetTable.Rows[$duplicateRow - 1]
                Add-ValidationWarningMessage -Report $report -MessageMap $report.WarningRows -RowNumber $duplicateRow -FieldName 'Description' -Message "ist doppelt: $($entry.Key)" -Row $row
            }
        }

        foreach ($entry in $macCounts.GetEnumerator()) {
            if ($entry.Value -le 1) { continue }

            foreach ($duplicateRow in $macRows[$entry.Key]) {
                $row = $targetTable.Rows[$duplicateRow - 1]
                Add-ValidationFieldMessage -Report $report -MessageMap $report.ErrorRows -RowNumber $duplicateRow -FieldName 'MACAddress' -Message "ist doppelt: $($entry.Key)" -Row $row
            }
        }

        foreach ($entry in $ipCounts.GetEnumerator()) {
            if ($entry.Value -le 1) { continue }

            foreach ($duplicateRow in $ipRows[$entry.Key]) {
                $row = $targetTable.Rows[$duplicateRow - 1]
                Add-ValidationFieldMessage -Report $report -MessageMap $report.ErrorRows -RowNumber $duplicateRow -FieldName 'IPAddress' -Message "ist doppelt: $($entry.Key)" -Row $row
            }
        }

        return $report
    }

    function Select-CsvFileByName {
        param([string]$FileName)

        if ([string]::IsNullOrWhiteSpace($FileName)) {
            $comboCsvFile.SelectedIndex = -1
            return $false
        }

        for ($index = 0; $index -lt $comboCsvFile.Items.Count; $index++) {
            if ([string]$comboCsvFile.Items[$index] -ieq $FileName) {
                $comboCsvFile.SelectedIndex = $index
                return $true
            }
        }

        return $false
    }

    function Update-CsvFileList {
        param([string]$PreferredFileName)

        $folderPath = Get-CsvFolderPath
        $previousSelection = [string]$comboCsvFile.SelectedItem
        $selectedFile = $null

        $files = @()
        if (-not [string]::IsNullOrWhiteSpace($folderPath) -and (Test-Path -LiteralPath $folderPath)) {
            $files = @(Get-ChildItem -LiteralPath $folderPath -Filter '*.csv' -File | Sort-Object Name)
        }

        $comboCsvFile.BeginUpdate()
        try {
            $comboCsvFile.Items.Clear()

            foreach ($file in $files) {
                [void]$comboCsvFile.Items.Add($file.Name)
            }

            if (-not [string]::IsNullOrWhiteSpace($PreferredFileName)) {
                if (Select-CsvFileByName -FileName $PreferredFileName) {
                    $selectedFile = [string]$comboCsvFile.SelectedItem
                } else {
                    $comboCsvFile.SelectedIndex = -1
                }
            } elseif (-not [string]::IsNullOrWhiteSpace($previousSelection)) {
                if (Select-CsvFileByName -FileName $previousSelection) {
                    $selectedFile = [string]$comboCsvFile.SelectedItem
                }
            }

            if ($null -eq $selectedFile -and [string]::IsNullOrWhiteSpace($PreferredFileName) -and $comboCsvFile.Items.Count -gt 0) {
                $comboCsvFile.SelectedIndex = 0
                $selectedFile = [string]$comboCsvFile.SelectedItem
            }

            if ($comboCsvFile.Items.Count -eq 0) {
                $comboCsvFile.SelectedIndex = -1
            }
        }
        finally {
            $comboCsvFile.EndUpdate()
        }

        return $selectedFile
    }

    function Set-AreaSelection {
        param([switch]$LoadCsv)

        $area = Get-SelectedAreaConfig
        if (-not $area) { return }

        $txtStartIp.Text = [string]$area.StartIp
        $txtEndIp.Text = [string]$area.EndIp

        $selectedFile = Update-CsvFileList -PreferredFileName ([string]$area.Csv)

        if ([string]::IsNullOrWhiteSpace($selectedFile)) {
            Set-StatusText "Vorlagen-CSV nicht gefunden: $($area.Csv)"
            return
        }

        if ($LoadCsv) {
            Import-SelectedCsv
        } else {
            Set-StatusText "Vorlage geladen: $([string]$comboArea.SelectedItem)"
        }
    }

    function Import-SelectedCsvIntoCurrentTable {
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = 'CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*'
        $openFileDialog.InitialDirectory = Get-CsvFolderPath
        $openFileDialog.Multiselect = $false

        try {
            if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                return
            }

            $importPath = $openFileDialog.FileName
            if ([string]::IsNullOrWhiteSpace($importPath)) {
                return
            }

            $grid.EndEdit()
            $bindingSource.EndEdit()

            $targetTable = @(Resolve-DataTableFromDataSource -DataSource $bindingSource.DataSource)[0]
            if (-not ($targetTable -is [System.Data.DataTable])) {
                $typeName = if ($null -eq $targetTable) { 'null' } else { $targetTable.GetType().FullName }
                throw "Import nicht möglich: Aktuelle Tabelle ist keine DataTable, sondern: $typeName"
            }

            try {
                $summary = Add-MacCsvRowsToDataTable `
                    -Table $targetTable `
                    -ImportPath $importPath `
                    -StartIp $txtStartIp.Text `
                    -EndIp $txtEndIp.Text `
                    -Settings $Settings

                if ($grid.Columns['Status']) {
                    $grid.Columns['Status'].ReadOnly = $true
                }

                $bindingSource.ResetBindings($false)
                $grid.Refresh()
                Update-ValidationDisplay | Out-Null
                Update-DataGridRowNumbers -Grid $grid
                $grid.Invalidate()

                $skippedTotal = $summary.SkippedDuplicateMac + $summary.SkippedInvalidMac + $summary.SkippedNoFreeIp
                $warningText = if ($summary.Warnings -and $summary.Warnings.Count -gt 0) {
                    "`r`n`r`nWarnungen:`r`n$($summary.Warnings -join "`r`n")"
                } else {
                    ''
                }

                $message = @"
Import abgeschlossen.

Importiert: $($summary.Imported)
Uebersprungen wegen doppelter MAC: $($summary.SkippedDuplicateMac)
Uebersprungen wegen ungueltiger MAC: $($summary.SkippedInvalidMac)
Uebersprungen, keine freie IP: $($summary.SkippedNoFreeIp)
Uebersprungen gesamt: $skippedTotal

Quelle:
$importPath

Hinweis:
Die vorgeschlagenen IP-Adressen wurden nur in die Tabelle eingetragen.
Zum Uebernehmen bitte "CSV speichern" klicken.$warningText
"@

                Set-StatusText "Import abgeschlossen: $($summary.Imported) importiert, $skippedTotal uebersprungen. Bitte CSV speichern."
                [System.Windows.Forms.MessageBox]::Show($message, 'CSV importieren', 'OK', 'Information') | Out-Null
            }
            catch {
                $fullMessage = @"
Fehler beim CSV-Import:

$($_.Exception.Message)
"@

                [System.Windows.Forms.MessageBox]::Show($fullMessage, 'CSV importieren', 'OK', 'Error') | Out-Null
            }
        }
        finally {
            $openFileDialog.Dispose()
        }
    }

    function Import-SelectedCsv {
        $csvPath = Get-CurrentCsvPath
        if ([string]::IsNullOrWhiteSpace($csvPath)) {
            throw 'Keine CSV-Datei ausgewaehlt.'
        }

        $script:CurrentCsvPath = $csvPath

        try {
            $result = Import-MacCsvToDataTable -Path $csvPath -RequiredColumns $Settings.RequiredColumns
            if ($null -eq $result -or $null -eq $result.Table) {
                throw 'CSV-Import lieferte keine Tabelle.'
            }

            if (-not ($result.Table -is [System.Data.DataTable])) {
                throw "CSV-Import lieferte keine DataTable, sondern: $($result.Table.GetType().FullName)"
            }

            $bindingSource.DataSource = $result.Table
            $script:CurrentDelimiter = if ($result.Delimiter) { [char]$result.Delimiter } else { ';' }

            if ($grid.Columns['Status']) {
                $grid.Columns['Status'].ReadOnly = $true
            }

            Update-ValidationDisplay | Out-Null
            Update-DataGridRowNumbers -Grid $grid
            $grid.Invalidate()
            Set-StatusText "CSV geladen: $csvPath | Neue Zeile unten eintragen, Zeile löschen mit Strg+Entf."
        }
        catch {
            $line = $_.InvocationInfo.ScriptLineNumber
            $scriptName = $_.InvocationInfo.ScriptName
            $codeLine = $_.InvocationInfo.Line
            $message = $_.Exception.Message

            $fullMessage = @"
Fehler beim CSV-Laden:

Datei:
$scriptName

Zeile:
$line

Code:
$codeLine

Meldung:
$message
"@

            try {
                $fallbackColumns = @('MACAddress','Description','Description2','IPAddress','Status')
                $fallbackTable = New-MacListDataTable -Columns $fallbackColumns

                if ($fallbackTable -is [System.Data.DataTable]) {
                    $bindingSource.DataSource = $fallbackTable
                }
            }
            catch {
                [void]$_.Exception
                # Falls sogar die Fallback-Tabelle fehlschlaegt, nichts weiter tun
            }

            Set-StatusText "Fehler beim Laden in Zeile ${line}: $message"
            [System.Windows.Forms.MessageBox]::Show($fullMessage, 'CSV laden', 'OK', 'Error') | Out-Null
        }
    }

    function Save-CurrentCsv {
        $csvPath = Get-CurrentCsvPath
        if ([string]::IsNullOrWhiteSpace($csvPath)) {
            throw 'Keine CSV-Datei ausgewaehlt.'
        }

        $form.Validate()
        $grid.EndEdit()
        $bindingSource.EndEdit()

        $targetTable = @(Resolve-DataTableFromDataSource -DataSource $bindingSource.DataSource)[0]
        if (-not ($targetTable -is [System.Data.DataTable])) {
            $typeName = if ($null -eq $targetTable) { 'null' } else { $targetTable.GetType().FullName }
            throw "Speichern nicht möglich: DataSource ist keine DataTable, sondern: $typeName"
        }

        $script:CurrentCsvPath = $csvPath
        Export-DataTableToMacCsv -Table $targetTable -Path $csvPath -Delimiter $script:CurrentDelimiter
        Set-StatusText "CSV gespeichert: $csvPath"
    }

    <#
    .SYNOPSIS
        Liefert die aktuellen Validierungsfehler als Textliste.

    .PARAMETER Table
        Die zu validierende DataSource oder DataTable.

    .PARAMETER StartIpText
        Start-IP aus der GUI.

    .PARAMETER EndIpText
        End-IP aus der GUI.

    .OUTPUTS
        System.String[]
    #>
    function Get-ValidationErrors {
        param(
            [Parameter(Mandatory)] [object]$Table,
            [Parameter(Mandatory)] [string]$StartIpText,
            [Parameter(Mandatory)] [string]$EndIpText
        )

        return @(Get-ValidationReport -Table $Table -StartIpText $StartIpText -EndIpText $EndIpText).Errors.ToArray()
    }

    function Update-ValidationDisplay {
        $report = Get-ValidationReport -Table $bindingSource.DataSource -StartIpText ([string]$txtStartIp.Text.Trim()) -EndIpText ([string]$txtEndIp.Text.Trim())

        foreach ($row in $grid.Rows) {
            if ($row.IsNewRow) { continue }

            $rowNumber = $row.Index + 1
            $hasErrors = $report.ErrorRows.ContainsKey($rowNumber)
            $hasWarnings = $report.WarningRows.ContainsKey($rowNumber)

            if ($hasErrors) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
                $row.ErrorText = ($report.ErrorRows[$rowNumber] -join [Environment]::NewLine)
            } elseif ($hasWarnings) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
                $row.ErrorText = ($report.WarningRows[$rowNumber] -join [Environment]::NewLine)
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.SystemColors]::Window
                $row.ErrorText = ''
            }
        }

        return $report
    }

    <#
    .SYNOPSIS
        Schreibt sichtbare Zeilennummern in die RowHeader-Zellen.

    .DESCRIPTION
        Die Nummerierung ist rein visuell und erzeugt keine CSV-Spalte.
    #>
    function Update-DataGridRowNumbers {
        param(
            [Parameter(Mandatory)]
            [System.Windows.Forms.DataGridView]$Grid
        )

        for ($i = 0; $i -lt $Grid.Rows.Count; $i++) {
            $row = $Grid.Rows[$i]

            if ($row.IsNewRow) {
                $row.HeaderCell.Value = ''
            } else {
                $row.HeaderCell.Value = ($i + 1).ToString()
            }
        }

        $Grid.RowHeadersVisible = $true
        $Grid.RowHeadersWidth = 55
        $Grid.Invalidate()
    }

    foreach ($areaName in (Get-OrderedAreaNames)) {
        [void]$comboArea.Items.Add($areaName)
    }

    if ($comboArea.Items.Count -gt 0) {
        $preferred = 'Screens und Beamer'
        if ($comboArea.Items.Contains($preferred)) {
            $comboArea.SelectedItem = $preferred
        } else {
            $comboArea.SelectedIndex = 0
        }
    }

    $comboArea.Add_SelectedIndexChanged({ Set-AreaSelection -LoadCsv })
    $btnBrowseFolder.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = 'CSV-Ordner auswählen'
        $dialog.SelectedPath = Get-CsvFolderPath

        try {
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtCsvFolder.Text = $dialog.SelectedPath
                Update-CsvFileList -PreferredFileName (Get-SelectedAreaCsvFileName) | Out-Null
                Set-StatusText "CSV-Ordner gewählt: $($txtCsvFolder.Text)"
            }
        }
        finally {
            $dialog.Dispose()
        }
    })

    $txtCsvFolder.Add_Leave({
        Update-CsvFileList -PreferredFileName (Get-SelectedAreaCsvFileName) | Out-Null
    })

    $btnLoad.Add_Click({ Import-SelectedCsvIntoCurrentTable })

    $btnSave.Add_Click({
        try {
            $grid.EndEdit()
            $bindingSource.EndEdit()

            $table = @(Resolve-DataTableFromDataSource -DataSource $bindingSource.DataSource)[0]
            $validation = Get-ValidationReport -Table $table -StartIpText ([string]$txtStartIp.Text.Trim()) -EndIpText ([string]$txtEndIp.Text.Trim())
            Update-ValidationDisplay | Out-Null

            if ($validation.Errors.Count -gt 0) {
                $errorText = Format-ValidationMessageBoxText -Headline 'Die Tabelle enthält Fehler. Bitte korrigieren Sie die folgenden Einträge:' -Messages @($validation.Errors) -ItemLabel 'Fehler'
                [System.Windows.Forms.MessageBox]::Show($errorText, 'CSV speichern', 'OK', 'Error') | Out-Null
                return
            }

            if ($validation.Warnings.Count -gt 0) {
                $warningText = Format-ValidationMessageBoxText -Headline 'Die Tabelle enthält Warnungen. Bitte prüfen Sie die folgenden Einträge:' -Messages @($validation.Warnings) -ItemLabel 'Warnungen'
                [System.Windows.Forms.MessageBox]::Show($warningText, 'CSV speichern', 'OK', 'Warning') | Out-Null
            }

            Save-CurrentCsv
        }
        catch {
            $fullMessage = @"
Fehler beim Speichern:

$($_.Exception.Message)
"@

            [System.Windows.Forms.MessageBox]::Show($fullMessage, 'CSV speichern', 'OK', 'Error') | Out-Null
        }
    })

    $btnRun.Add_Click({
        try {
            $grid.EndEdit()
            $bindingSource.EndEdit()

            $table = @(Resolve-DataTableFromDataSource -DataSource $bindingSource.DataSource)[0]
            $validation = Get-ValidationReport -Table $table -StartIpText ([string]$txtStartIp.Text.Trim()) -EndIpText ([string]$txtEndIp.Text.Trim())
            Update-ValidationDisplay | Out-Null

            if ($validation.Errors.Count -gt 0) {
                $errorText = Format-ValidationMessageBoxText -Headline 'Die Tabelle enthält Fehler. Bitte korrigieren Sie die folgenden Einträge:' -Messages @($validation.Errors) -ItemLabel 'Fehler'
                [System.Windows.Forms.MessageBox]::Show($errorText, 'Validierung', 'OK', 'Warning') | Out-Null
                return
            }

            if ($validation.Warnings.Count -gt 0) {
                $warningText = Format-ValidationMessageBoxText -Headline 'Die Tabelle enthält Warnungen. Bitte prüfen Sie die folgenden Einträge:' -Messages @($validation.Warnings) -ItemLabel 'Warnungen'
                [System.Windows.Forms.MessageBox]::Show($warningText, 'Validierung', 'OK', 'Warning') | Out-Null
            }

            $script:CurrentCsvPath = Get-CurrentCsvPath
            if ([string]::IsNullOrWhiteSpace($script:CurrentCsvPath)) {
                throw 'Keine CSV-Datei ausgewaehlt.'
            }

            $backupPath = Backup-MacCsv -Path $script:CurrentCsvPath
            Save-CurrentCsv

            $btnRun.Enabled = $false
            Set-StatusText 'Reservierungen werden ausgefuehrt ...'

            $summary = Invoke-DhcpMacReservations -Table $table -Settings $Settings -StatusCallback ${function:Set-StatusText}

            Save-CurrentCsv

            $message = "Fertig.`r`nVerarbeitet: $($summary.Processed)`r`nFehler: $($summary.Errors)`r`nDHCP-Server: $($summary.DhcpServer)"
            if ($backupPath) { $message += "`r`nBackup: $backupPath" }

            Set-StatusText $message.Replace("`r`n", ' | ')
            [System.Windows.Forms.MessageBox]::Show($message, 'DHCP MAC-Reservierung', 'OK', 'Information') | Out-Null
        }
        catch {
            Set-StatusText "Fehler: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Fehler', 'OK', 'Error') | Out-Null
        }
        finally {
            $btnRun.Enabled = $true
        }
    })

    $btnClose.Add_Click({ $form.Close() })

    if ($comboArea.Items.Count -gt 0) {
        Set-AreaSelection -LoadCsv
    } else {
        Update-CsvFileList -PreferredFileName (Get-SelectedAreaCsvFileName) | Out-Null
    }

    [void]$form.ShowDialog()
}
