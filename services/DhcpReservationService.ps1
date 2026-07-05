<#
.SYNOPSIS
    DHCP-Reservierungen und Filter fuer MAC-Listen.

.DESCRIPTION
    Fuehrt Reservierungen aus, prueft/aktualisiert vorhandene DHCP-Eintraege und
    setzt Statuswerte in der CSV. Der Status dient als Protokoll pro Eintrag.

.NOTES
    Projekt: DHCP MAC-Reservierung
    Umgebung: Windows Server 2022 / Windows PowerShell 5.1
#>

Import-Module DhcpServer -ErrorAction SilentlyContinue

function Resolve-DhcpServerName {
    param([hashtable]$Settings)

    $candidates = New-Object System.Collections.Generic.List[string]

    if ($Settings -and $Settings.ContainsKey('DhcpServer')) {
        $configured = [string]$Settings.DhcpServer
        if (-not [string]::IsNullOrWhiteSpace($configured)) {
            $candidates.Add($configured.Trim())
        }
    }

    try {
        $serversInDc = @(Get-DhcpServerInDC -ErrorAction Stop)
        foreach ($server in $serversInDc) {
            if ($server -and -not [string]::IsNullOrWhiteSpace($server.DnsName)) {
                $candidates.Add(([string]$server.DnsName).Trim())
            }
        }
    } catch {
        Write-Warning "Get-DhcpServerInDC fehlgeschlagen: $($_.Exception.Message)"
    }

    if (-not [string]::IsNullOrWhiteSpace($env:COMPUTERNAME)) {
        $candidates.Add($env:COMPUTERNAME.Trim())
    }

    foreach ($candidate in $candidates | Select-Object -Unique) {
        try {
            Get-DhcpServerv4Scope -ComputerName $candidate -ErrorAction Stop | Out-Null
            return $candidate
        } catch {
            Write-Warning "DHCP-Server '$candidate' ist nicht verwendbar: $($_.Exception.Message)"
        }
    }

    return $null
}

function Add-DhcpV4ReservationSafe {
    param(
        [Parameter(Mandatory)] [string]$ComputerName,
        [Parameter(Mandatory)] [System.Net.IPAddress]$ScopeId,
        [Parameter(Mandatory)] [System.Net.IPAddress]$IPAddress,
        [Parameter(Mandatory)] [string]$ClientId,
        [string]$Name,
        [string]$Description
    )

    $params = @{
        ComputerName = $ComputerName
        ScopeId      = $ScopeId
        IPAddress    = $IPAddress
        ClientId     = [string]$ClientId
    }

    if ($Name)        { $params['Name']        = $Name }
    if ($Description) { $params['Description'] = $Description }

    Add-DhcpServerv4Reservation @params
}

function Convert-IPv4AddressToUInt32 {
    param([Parameter(Mandatory)] [System.Net.IPAddress]$IPAddress)

    $bytes = $IPAddress.GetAddressBytes()
    if ($bytes.Length -ne 4) {
        throw 'Nur IPv4-Adressen werden unterstuetzt.'
    }

    if ([BitConverter]::IsLittleEndian) {
        [Array]::Reverse($bytes)
    }

    return [BitConverter]::ToUInt32($bytes, 0)
}

function Convert-UInt32ToIPv4String {
    param([Parameter(Mandatory)] [uint32]$Value)

    $bytes = [BitConverter]::GetBytes($Value)
    if ([BitConverter]::IsLittleEndian) {
        [Array]::Reverse($bytes)
    }

    return ('{0}.{1}.{2}.{3}' -f $bytes[0], $bytes[1], $bytes[2], $bytes[3])
}

function Get-NextFreeIpInRange {
    <#
    .SYNOPSIS
        Ermittelt die naechste freie IPv4-Adresse in einem Bereich.

    .DESCRIPTION
        Der Bereich stammt aus den organisatorischen Areas und begrenzt nur
        Vorschlaege bzw. Validierung. Der DHCP-Scope selbst bleibt global.

    .PARAMETER StartIp
        Untere IP-Grenze des vorgeschlagenen Bereichs.

    .PARAMETER EndIp
        Obere IP-Grenze des vorgeschlagenen Bereichs.

    .PARAMETER UsedIps
        Sammlung bereits belegter IPv4-Adressen.

    .OUTPUTS
        System.String
    #>
    param(
        [Parameter(Mandatory)] [string]$StartIp,
        [Parameter(Mandatory)] [string]$EndIp,
        [Parameter(Mandatory)] $UsedIps
    )

    $startIpAddress = $null
    if (-not [System.Net.IPAddress]::TryParse($StartIp, [ref]$startIpAddress)) {
        throw "Ungueltige Start-IP: $StartIp"
    }

    $endIpAddress = $null
    if (-not [System.Net.IPAddress]::TryParse($EndIp, [ref]$endIpAddress)) {
        throw "Ungueltige End-IP: $EndIp"
    }

    if ($startIpAddress.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork -or
        $endIpAddress.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
        throw 'Nur IPv4-Adressen werden unterstuetzt.'
    }

    $startValue = Convert-IPv4AddressToUInt32 -IPAddress $startIpAddress
    $endValue = Convert-IPv4AddressToUInt32 -IPAddress $endIpAddress

    if ($startValue -gt $endValue) {
        throw 'Start-IP muss kleiner oder gleich End-IP sein.'
    }

    for ($value = $startValue; $value -le $endValue; $value++) {
        $candidate = Convert-UInt32ToIPv4String -Value ([uint32]$value)

        $isUsed = $false
        if ($null -ne $UsedIps) {
            try {
                $isUsed = [bool]$UsedIps.Contains($candidate)
            }
            catch {
                $isUsed = $UsedIps -contains $candidate
            }
        }

        if (-not $isUsed) {
            return $candidate
        }
    }

    return $null
}

function Set-DhcpV4ReservationSafe {
    param(
        [Parameter(Mandatory)] [string]$ComputerName,
        [Parameter(Mandatory)] [System.Net.IPAddress]$IPAddress,
        [Parameter(Mandatory)] [string]$ClientId,
        [string]$Name,
        [string]$Description
    )

    $params = @{
        ComputerName = $ComputerName
        IPAddress    = $IPAddress
        ClientId     = [string]$ClientId
    }

    if ($Name)        { $params['Name']        = $Name }
    if ($Description) { $params['Description'] = $Description }

    Set-DhcpServerv4Reservation @params
}

function Invoke-DhcpMacReservations {
    <#
    .SYNOPSIS
        Fuehrt DHCP-Reservierungen fuer alle gueltigen Tabellenzeilen aus.

    .DESCRIPTION
        Schreibt Statuswerte im Format LEVEL | yyyy-MM-dd HH:mm:ss | Meldung in die
        Status-Spalte der Arbeits-CSV. Statuswerte sind das einzige Protokoll pro Zeile.

        Der DHCP-Scope wird global aus Settings.DhcpScope verwendet, weil das Projekt
        mit einem gemeinsamen /16-Scope arbeitet. Areas sind nur organisatorische
        Start-/Endbereiche fuer Vorschlaege und Validierung.

    .PARAMETER Table
        Tabelle oder normale DataSource mit MAC- und IP-Daten.

    .PARAMETER Settings
        Projektweite Einstellungen inkl. globalem DHCP-Scope.

    .PARAMETER StatusCallback
        Rueckruffunktion fuer Statusmeldungen im GUI.

    .OUTPUTS
        PSCustomObject mit Processed, Errors und DhcpServer.
    #>
    param(
        [Parameter(Mandatory)] [object]$Table,
        [Parameter(Mandatory)] [hashtable]$Settings,
        [Parameter(Mandatory)] [scriptblock]$StatusCallback
    )

    $targetTable = @(Resolve-DataTableFromDataSource -DataSource $Table)[0]
    if (-not ($targetTable -is [System.Data.DataTable])) {
        throw "Table ist keine DataTable, sondern: $($targetTable.GetType().FullName)"
    }

    $processed = 0
    $errors = 0

    $scopeId = [System.Net.IPAddress]$Settings.DhcpScope
    $scopeIdStr = $scopeId.ToString()

    $dhcpServer = Resolve-DhcpServerName -Settings $Settings
    if ([string]::IsNullOrWhiteSpace($dhcpServer)) {
        throw 'Es konnte kein verwendbarer DHCP-Server ermittelt werden.'
    }

    & $StatusCallback "Arbeite mit DHCP-Server $dhcpServer und Scope $scopeIdStr ..."

    $reservationsAll = @(Get-DhcpServerv4Reservation -ComputerName $dhcpServer -ScopeId $scopeIdStr -ErrorAction SilentlyContinue)

    foreach ($row in $targetTable.Rows) {
        if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }

        $name   = ([string]$row['Description']).Trim()
        $descr  = ([string]$row['Description2']).Trim()
        $ipStr  = ([string]$row['IPAddress']).Trim()
        $macRaw = ([string]$row['MACAddress']).Trim()

        $macNorm = ConvertTo-NormalizedMac -Mac $macRaw
        $macHyphen = Format-MacHyphen -Mac $macRaw

        if (-not $macNorm -or $macNorm.Length -ne 12) {
            $row['Status'] = New-MacListStatus -Level 'FEHLER' -Message "MAC-Adresse ungültig: $macRaw"
            $errors++
            continue
        }

        $row['MACAddress'] = $macHyphen

        try {
            $ipObj = $null
            if (-not [System.Net.IPAddress]::TryParse($ipStr, [ref]$ipObj)) {
                throw "Ungueltige IP in CSV: $ipStr ($name)"
            }

            $ip = [System.Net.IPAddress]$ipStr
            $ipStrCanon = $ip.ToString()

            $existingMACReservation = $reservationsAll | Where-Object {
                $_.ClientId -and ((($_.ClientId) -replace '[^0-9A-Fa-f]', '').ToUpper() -eq $macNorm)
            }

            $existingIPReservation = $reservationsAll | Where-Object {
                $_.IPAddress -and ($_.IPAddress.ToString() -eq $ipStrCanon)
            }

            if ($existingMACReservation) {
                $assignedIP = $existingMACReservation.IPAddress
                $actionMsg = "Bereits reserviert: MAC $macHyphen -> IP $assignedIP"

                $needsUpdate = $false
                if ($name -and ($existingMACReservation.Name -ne $name)) { $needsUpdate = $true }
                if ($descr -and ($existingMACReservation.Description -ne $descr)) { $needsUpdate = $true }

                if ($needsUpdate) {
                    Set-DhcpV4ReservationSafe -ComputerName $dhcpServer -IPAddress $assignedIP -ClientId $macHyphen -Name $name -Description $descr
                    $reservationsAll = @(Get-DhcpServerv4Reservation -ComputerName $dhcpServer -ScopeId $scopeIdStr -ErrorAction SilentlyContinue)
                    $actionMsg = "Reservierung aktualisiert: $name -> $assignedIP / $macHyphen"
                }

                if ($assignedIP.ToString() -ne $ipStrCanon) {
                    $actionMsg += " | Hinweis: CSV-IP $ipStrCanon ungleich reservierte IP $assignedIP."
                }
            }
            elseif ($existingIPReservation) {
                $assignedMAC = $existingIPReservation.ClientId
                $actionMsg = "Uebersprungen: IP $ipStrCanon ist bereits reserviert durch $assignedMAC"
            }
            else {
                # Wichtig: Nicht ScopeId und IPAddress gemeinsam an Get-DhcpServerv4Lease uebergeben.
                $lease = Get-DhcpServerv4Lease -ComputerName $dhcpServer -IPAddress $ip -ErrorAction SilentlyContinue
                if ($lease) {
                    & $StatusCallback "Hinweis: IP $ipStrCanon ist aktuell geleast durch $($lease.ClientId)."
                }

                Add-DhcpV4ReservationSafe -ComputerName $dhcpServer -ScopeId $scopeId -IPAddress $ip -ClientId $macHyphen -Name $name -Description $descr
                $reservationsAll = @(Get-DhcpServerv4Reservation -ComputerName $dhcpServer -ScopeId $scopeIdStr -ErrorAction SilentlyContinue)
                $actionMsg = "Reservierung erstellt: $name -> $ipStrCanon / $macHyphen"
            }

            $filterMsg = Add-DhcpMacFilterIfMissing -Settings $Settings -MacAddress $macHyphen -MacNorm $macNorm -Description $name

            $row['Status'] = New-MacListStatus -Level 'OK' -Message "$actionMsg | $filterMsg"
            $processed++
        }
        catch {
            $row['Status'] = New-MacListStatus -Level 'FEHLER' -Message $_.Exception.Message
            $errors++
        }
    }

    return [PSCustomObject]@{
        Processed = $processed
        Errors    = $errors
        DhcpServer = $dhcpServer
    }
}

function Add-DhcpMacFilterIfMissing {
    <#
    .SYNOPSIS
        Stellt sicher, dass ein MAC-Filter vorhanden ist.

    .OUTPUTS
        System.String
    #>
    param(
        [Parameter(Mandatory)] [hashtable]$Settings,
        [Parameter(Mandatory)] [string]$MacAddress,
        [Parameter(Mandatory)] [string]$MacNorm,
        [string]$Description
    )

    $filterServer = [string]$Settings.FilterServer
    $filterListName = [string]$Settings.FilterListName

    if ([string]::IsNullOrWhiteSpace($filterServer) -or [string]::IsNullOrWhiteSpace($filterListName)) {
        return 'Filter uebersprungen: FilterServer oder FilterListName nicht gesetzt.'
    }

    try {
        $existingFilter = Get-DhcpServerv4Filter -ComputerName $filterServer -List $filterListName -ErrorAction SilentlyContinue |
            Where-Object { (ConvertTo-NormalizedMac -Mac $_.MacAddress) -eq $MacNorm }

        if (-not $existingFilter) {
            $filterParams = @{
                ComputerName = $filterServer
                List         = $filterListName
                MacAddress   = $MacAddress
            }

            if ($Description) { $filterParams['Description'] = $Description }

            Add-DhcpServerv4Filter @filterParams
            return "Filter hinzugefuegt ($filterListName@$filterServer)."
        }

        return "Filter bereits vorhanden ($filterListName@$filterServer)."
    }
    catch {
        return "Filter-Fehler: $($_.Exception.Message)"
    }
}
