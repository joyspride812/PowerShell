$Servers = Get-ADObject `
                        -Filter { (objectCategory -eq 'msExchExchangeServer') -and (objectClass    -eq 'msExchExchangeServer')   } `
                        -SearchBase (Get-ADRootDSE).configurationNamingContext -Properties * |
                            where {($_.msExchCurrentServerRoles -band 54) -eq 54} | select  -ExpandProperty networkAddress |
                            where {$_ -match "ncacn_ip_tcp"} | %{$_ -replace 'ncacn_ip_tcp:',""} |
                            foreach { 
                                Test-Connection $_ -ErrorAction SilentlyContinue  -Count 1
                            }
$bestServer = $Servers | sort ResponseTime | select -First 1 -Property Address
$E2k13PsSession = New-PSSession  `
                    -ConfigurationName  Microsoft.Exchange `
                    -ConnectionUri     "http://$($bestServer.Address)/powershell/"`
                    -Authentication     Kerberos `
                    -WarningAction      SilentlyContinue `
                    -ErrorAction        SilentlyContinue
$null = Import-PSSession $E2k13PsSession
"*********","**********","**********","**********","************","**********","*********","**********"|%{Get-MessageTrackingLog -Server $_ -sender "ordercd@ixora-auto.ru" -recipients "OZH_AVTOPRIME@rolf.ru" -Start "2023/07/04 10:00:00" -End "2023/07/05 14:00:00"|select timestamp,recipients,eventid,messagesubject,RecipientStatus,MessageId}

Remove-PSSession $E2k13PsSession


  