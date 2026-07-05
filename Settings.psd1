<#
.SYNOPSIS
    Projektkonfiguration fuer die DHCP MAC-Reservierung.

.DESCRIPTION
    Definiert den globalen DHCP-Scope, die Arbeits-CSV je Bereich und die
    organisatorischen Areas fuer Vorschlaege und Validierung.

.NOTES
    Projekt: DHCP MAC-Reservierung
    Umgebung: Windows Server 2022 / Windows PowerShell 5.1
#>

@{
    # Arbeitsordner mit den MAC-CSV-Dateien.
    MacListenPfad = 'D:\System\MAC-Listen\'

    # Standard-DHCP-Server. Leer lassen, wenn automatisch ermittelt werden soll.
    DhcpServer = 'srv-wbg-1'

    # Globaler DHCP-Scope 172.16.0.0 fuer das gesamte /16-Netz.
    # Dieser Scope ist global und wird nicht aus den Areas bezogen.
    DhcpScope = '172.16.0.0'

    # DHCP-MAC-Filter
    FilterServer   = 'srv-wbg-1'
    FilterListName = 'Allow' # Allow oder Deny

    # CSV-Spalten, die das Tool erwartet
    RequiredColumns = @(
        'MACAddress'
        'Description'
        'Description2'
        'IPAddress'
    )

    # Areas sind nur organisatorische Teilbereiche innerhalb des globalen Scopes.
    # StartIp/EndIp begrenzen nur Vorschlaege und Validierung.
    # Csv ist die jeweilige Arbeitsdatei des Bereichs.
    Areas = @{
        'Server' = @{
            Order   = 10
            Csv     = 'mac-server.csv'
            StartIp = '172.16.0.1'
            EndIp   = '172.16.0.19'
        }

        'System und Drucker' = @{
            Order   = 20
            Csv     = 'mac-system-drucker.csv'
            StartIp = '172.16.0.20'
            EndIp   = '172.16.0.99'
        }

        'Screens und Beamer' = @{
            Order   = 30
            Csv     = 'mac-screens.csv'
            StartIp = '172.16.0.100'
            EndIp   = '172.16.0.254'
        }

        'EDV-Raeume' = @{
            Order   = 40
            Csv     = 'mac-edv.csv'
            StartIp = '172.16.1.1'
            EndIp   = '172.16.1.254'
        }

        'IPad-Koffer' = @{
            Order   = 50
            Csv     = 'mac-ipadkoffer.csv'
            StartIp = '172.16.2.1'
            EndIp   = '172.16.2.254'
        }

        'Sonstiger Bestand' = @{
            Order   = 60
            Csv     = 'mac-bestand.csv'
            StartIp = '172.16.3.1'
            EndIp   = '172.16.3.254'
        }

        'LehrerDienstgeraete' = @{
            Order   = 70
            Csv     = 'mac-ldg.csv'
            StartIp = '172.16.20.1'
            EndIp   = '172.16.20.254'
        }

        'LehrerPrivatGeraete' = @{
            Order   = 80
            Csv     = 'mac-lpg.csv'
            StartIp = '172.16.21.1'
            EndIp   = '172.16.21.254'
        }
    }
}
