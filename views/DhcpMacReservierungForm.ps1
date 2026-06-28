<#
.SYNOPSIS
    WinForms-Oberflaeche fuer DHCP MAC-Reservierung.
#>

function Show-DhcpMacReservierungForm {
    param([Parameter(Mandatory)] [hashtable]$Settings)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'DHCP MAC-Reservierung'
    $form.Size = New-Object System.Drawing.Size(980,640)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(900,560)

    $fontDefault = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.Font = $fontDefault

    $lblArea = New-Object System.Windows.Forms.Label
    $lblArea.Location = New-Object System.Drawing.Point(12,18)
    $lblArea.Size = New-Object System.Drawing.Size(95,22)
    $lblArea.Text = 'Adressbereich:'
    $form.Controls.Add($lblArea)

    $comboArea = New-Object System.Windows.Forms.ComboBox
    $comboArea.Location = New-Object System.Drawing.Point(110,15)
    $comboArea.Size = New-Object System.Drawing.Size(300,24)
    $comboArea.DropDownStyle = 'DropDownList'
    $form.Controls.Add($comboArea)

    $lblRange = New-Object System.Windows.Forms.Label
    $lblRange.Location = New-Object System.Drawing.Point(430,18)
    $lblRange.Size = New-Object System.Drawing.Size(520,22)
    $lblRange.Text = 'Bereich:'
    $form.Controls.Add($lblRange)

    $lblCsv = New-Object System.Windows.Forms.Label
    $lblCsv.Location = New-Object System.Drawing.Point(12,48)
    $lblCsv.Size = New-Object System.Drawing.Size(930,22)
    $lblCsv.Text = 'CSV-Datei:'
    $form.Controls.Add($lblCsv)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12,82)
    $grid.Size = New-Object System.Drawing.Size(940,400)
    $grid.Anchor = 'Top, Bottom, Left, Right'
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.AllowUserToAddRows = $true
    $grid.AllowUserToDeleteRows = $true
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.RowHeadersVisible = $false
    $grid.EditMode = 'EditOnKeystrokeOrF2'
    $form.Controls.Add($grid)

    $bindingSource = New-Object System.Windows.Forms.BindingSource
    $grid.DataSource = $bindingSource

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Location = New-Object System.Drawing.Point(12,500)
    $btnLoad.Size = New-Object System.Drawing.Size(110,30)
    $btnLoad.Anchor = 'Bottom, Left'
    $btnLoad.Text = 'CSV laden'
    $form.Controls.Add($btnLoad)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(132,500)
    $btnSave.Size = New-Object System.Drawing.Size(120,30)
    $btnSave.Anchor = 'Bottom, Left'
    $btnSave.Text = 'CSV speichern'
    $form.Controls.Add($btnSave)

    $btnAddRow = New-Object System.Windows.Forms.Button
    $btnAddRow.Location = New-Object System.Drawing.Point(262,500)
    $btnAddRow.Size = New-Object System.Drawing.Size(110,30)
    $btnAddRow.Anchor = 'Bottom, Left'
    $btnAddRow.Text = 'Zeile +'
    $form.Controls.Add($btnAddRow)

    $btnDeleteRow = New-Object System.Windows.Forms.Button
    $btnDeleteRow.Location = New-Object System.Drawing.Point(382,500)
    $btnDeleteRow.Size = New-Object System.Drawing.Size(120,30)
    $btnDeleteRow.Anchor = 'Bottom, Left'
    $btnDeleteRow.Text = 'Zeile loeschen'
    $form.Controls.Add($btnDeleteRow)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(690,500)
    $btnRun.Size = New-Object System.Drawing.Size(170,30)
    $btnRun.Anchor = 'Bottom, Right'
    $btnRun.Text = 'Reservierungen ausfuehren'
    $form.Controls.Add($btnRun)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(870,500)
    $btnClose.Size = New-Object System.Drawing.Size(82,30)
    $btnClose.Anchor = 'Bottom, Right'
    $btnClose.Text = 'Schliessen'
    $form.Controls.Add($btnClose)

    $status = New-Object System.Windows.Forms.Label
    $status.Location = New-Object System.Drawing.Point(12,545)
    $status.Size = New-Object System.Drawing.Size(940,35)
    $status.Anchor = 'Bottom, Left, Right'
    $status.Text = 'Bereit.'
    $form.Controls.Add($status)

    $script:CurrentCsvPath = $null
    $script:CurrentDelimiter = ';'
    $script:CurrentAreaConfig = $null

    function Set-StatusText {
        param([string]$Text)
        $status.Text = $Text
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Get-SelectedAreaConfig {
        if (-not $comboArea.SelectedItem) { return $null }
        return $Settings.Areas[[string]$comboArea.SelectedItem]
    }

    function Load-SelectedCsv {
        $area = Get-SelectedAreaConfig
        if (-not $area) { return }

        $script:CurrentAreaConfig = $area
        $script:CurrentCsvPath = Join-Path ([string]$Settings.MacListenPfad) ([string]$area.Csv)

        $lblRange.Text = "Bereich: $($area.Range) | Scope: $($area.Scope)"
        $lblCsv.Text = "CSV-Datei: $script:CurrentCsvPath"

        try {
            $result = Import-MacCsvToDataTable -Path $script:CurrentCsvPath -RequiredColumns $Settings.RequiredColumns
            if ($null -eq $result -or $null -eq $result.Table) {
                throw 'CSV-Import lieferte keine Tabelle.'
            }

            if (-not ($result.Table -is [System.Data.DataTable])) {
                throw "CSV-Import lieferte keine DataTable, sondern: $($result.Table.GetType().FullName)"
            }

            $bindingSource.DataSource = $result.Table
            $script:CurrentDelimiter = [char]$result.Delimiter

            if ($grid.Columns['Status']) {
                $grid.Columns['Status'].ReadOnly = $true
            }

            Set-StatusText "CSV geladen: $script:CurrentCsvPath"
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
    } catch {
        # Falls sogar die Fallback-Tabelle fehlschlaegt, nichts weiter tun
    }

    Set-StatusText "Fehler beim Laden in Zeile ${line}: $message"
    [System.Windows.Forms.MessageBox]::Show($fullMessage, 'CSV laden', 'OK', 'Error') | Out-Null
}
    }

    function Save-CurrentCsv {
        if ([string]::IsNullOrWhiteSpace($script:CurrentCsvPath)) {
            throw 'Keine CSV-Datei ausgewaehlt.'
        }

        $grid.EndEdit()
        $bindingSource.EndEdit()

        $table = [System.Data.DataTable]$bindingSource.DataSource
        Export-DataTableToMacCsv -Table $table -Path $script:CurrentCsvPath -Delimiter $script:CurrentDelimiter
        Set-StatusText "CSV gespeichert: $script:CurrentCsvPath"
    }

    foreach ($areaName in ($Settings.Areas.Keys | Sort-Object)) {
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

    $comboArea.Add_SelectedIndexChanged({ Load-SelectedCsv })
    $btnLoad.Add_Click({ Load-SelectedCsv })

    $btnSave.Add_Click({
        try {
            Save-CurrentCsv
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'CSV speichern', 'OK', 'Error') | Out-Null
        }
    })

    $btnAddRow.Add_Click({
        $table = [System.Data.DataTable]$bindingSource.DataSource
        if ($null -eq $table) { return }
        $newRow = $table.NewRow()
        [void]$table.Rows.Add($newRow)
    })

    $btnDeleteRow.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) { return }
        $confirm = [System.Windows.Forms.MessageBox]::Show('Ausgewaehlte Zeile wirklich loeschen?', 'Zeile loeschen', 'YesNo', 'Question')
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            $grid.Rows.RemoveAt($grid.SelectedRows[0].Index)
        }
    })

    $btnRun.Add_Click({
        try {
            $grid.EndEdit()
            $bindingSource.EndEdit()

            $table = [System.Data.DataTable]$bindingSource.DataSource
            if ($null -eq $table) { throw 'Keine CSV-Daten geladen.' }

            $validationErrors = Test-MacListDataTable -Table $table
            if ($validationErrors.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show(($validationErrors -join "`r`n"), 'Validierung', 'OK', 'Warning') | Out-Null
                return
            }

            $backupPath = Backup-MacCsv -Path $script:CurrentCsvPath
            Save-CurrentCsv

            $btnRun.Enabled = $false
            Set-StatusText 'Reservierungen werden ausgefuehrt ...'

            $summary = Invoke-DhcpMacReservations -Table $table -Settings $Settings -Scope ([string]$script:CurrentAreaConfig.Scope) -StatusCallback ${function:Set-StatusText}

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

    Load-SelectedCsv

    [void]$form.ShowDialog()
}
