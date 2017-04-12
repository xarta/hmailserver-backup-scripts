# Permissive for now ... eventually look at restricting to 25, 587, 993 for email
# TODO: pass these rules as arguments read from Xarta.json
Remove-NetFirewallRule -DisplayName "_XARTA EMAIL"
New-NetFirewallRule -DisplayName "_XARTA EMAIL" -Group "_XARTA" -Description "inbound ports for hMailServer: 25 smtp from incoming relays, 587 for smtp submissions from clients, 993 for Imap"-Direction Inbound -Action Allow -Profile Private,Domain,Public -EdgeTraversalPolicy Allow -Protocol TCP -LocalPort 25,110,143,465,585,587,993,995
Remove-NetFirewallRule -DisplayName "_XARTA RDP 8991-8998 TCP-User-IN"
New-NetFirewallRule -DisplayName "_XARTA RDP 8991-8998 TCP-User-IN" -Group "_XARTA" -Direction Inbound -Action Allow -Profile Private,Domain,Public -EdgeTraversalPolicy Allow -Program "%SystemRoot%\system32\svchost.exe" -Protocol TCP -LocalPort 8991-8998
Remove-NetFirewallRule -DisplayName "_XARTA RDP 8991-8998 UDP-User-IN"
New-NetFirewallRule -DisplayName "_XARTA RDP 8991-8998 UDP-User-IN" -Group "_XARTA" -Direction Inbound -Action Allow -Profile Private,Domain,Public -EdgeTraversalPolicy Allow -Program "%SystemRoot%\system32\svchost.exe" -Protocol UDP -LocalPort 8991-8998