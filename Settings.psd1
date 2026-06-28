@{
    # Ordner mit den MAC-CSV-Dateien
    MacListenPfad = 'D:\System\MAC-Listen\'

    # Standard-DHCP-Server. Leer lassen, wenn automatisch ermittelt werden soll.
    DhcpServer = 'srv-wbg-1'

    # DHCP-MAC-Filter
    FilterServer   = 'srv-wbg-1'
    FilterListName = 'Allow' # Allow oder Deny

    # Logging
    LogSubFolder = 'Logs'

    # CSV-Spalten, die das Tool erwartet
    RequiredColumns = @(
        'MACAddress'
        'Description'
        'Description2'
        'IPAddress'
    )

    Areas = @{
        'Server' = @{
            Csv   = 'mac-server.csv'
            Scope = '172.16.0.0'
            Range = '172.16.0.1 - 172.16.0.19'
        }

        'System und Drucker' = @{
            Csv   = 'mac-system-drucker.csv'
            Scope = '172.16.0.0'
            Range = '172.16.0.20 - 172.16.0.99'
        }

        'Screens und Beamer' = @{
            Csv   = 'mac-screens.csv'
            Scope = '172.16.0.0'
            Range = '172.16.0.100 - 172.16.0.254'
        }

        'EDV-Raeume' = @{
            Csv   = 'mac-edv.csv'
            Scope = '172.16.1.0'
            Range = '172.16.1.1 - 172.16.1.254'
        }

        'IPad-Koffer' = @{
            Csv   = 'mac-ipadkoffer.csv'
            Scope = '172.16.2.0'
            Range = '172.16.2.1 - 172.16.2.254'
        }

        'Sonstiger Bestand' = @{
            Csv   = 'mac-bestand.csv'
            Scope = '172.16.3.0'
            Range = '172.16.3.1 - 172.16.3.254'
        }

        'LehrerDienstgeraete' = @{
            Csv   = 'mac-ldg.csv'
            Scope = '172.16.20.0'
            Range = '172.16.20.1 - 172.16.20.254'
        }

        'LehrerPrivatGeraete' = @{
            Csv   = 'mac-lpg.csv'
            Scope = '172.16.21.0'
            Range = '172.16.21.1 - 172.16.21.254'
        }
    }
}
