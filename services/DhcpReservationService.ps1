<#
.SYNOPSIS
    DHCP-Service fuer Reservierungen und MAC-Filter.
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
    param(
        [Parameter(Mandatory)] [System.Data.DataTable]$Table,
        [Parameter(Mandatory)] [hashtable]$Settings,
        [Parameter(Mandatory)] [string]$Scope,
        [Parameter(Mandatory)] [scriptblock]$StatusCallback
    )

    $nowStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $processed = 0
    $errors = 0

    $scopeId = [System.Net.IPAddress]$Scope
    $scopeIdStr = $scopeId.ToString()

    $dhcpServer = Resolve-DhcpServerName -Settings $Settings
    if ([string]::IsNullOrWhiteSpace($dhcpServer)) {
        throw 'Es konnte kein verwendbarer DHCP-Server ermittelt werden.'
    }

    & $StatusCallback "Arbeite mit DHCP-Server $dhcpServer und Scope $scopeIdStr ..."

    $reservationsAll = @(Get-DhcpServerv4Reservation -ComputerName $dhcpServer -ScopeId $scopeIdStr -ErrorAction SilentlyContinue)

    foreach ($row in $Table.Rows) {
        if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }

        $name   = ([string]$row['Description']).Trim()
        $descr  = ([string]$row['Description2']).Trim()
        $ipStr  = ([string]$row['IPAddress']).Trim()
        $macRaw = ([string]$row['MACAddress']).Trim()

        $macNorm = Normalize-Mac -Mac $macRaw
        $macHyphen = Format-MacHyphen -Mac $macRaw

        if (-not $macNorm -or $macNorm.Length -ne 12) {
            $row['Status'] = "Fehler: Ungueltige MAC '$macRaw' | $nowStamp"
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

            $row['Status'] = "OK: $actionMsg | $filterMsg | $nowStamp"
            $processed++
        }
        catch {
            $row['Status'] = "Fehler: $($_.Exception.Message) | $nowStamp"
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
            Where-Object { (Normalize-Mac -Mac $_.MacAddress) -eq $MacNorm }

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
