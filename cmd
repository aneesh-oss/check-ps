New-NetFirewallRule -DisplayName "Allow New Relic Download" `
    -Direction Outbound -Action Allow -RemoteAddress "download.newrelic.com"

Resolve-DnsName download.newrelic.com

Test-NetConnection download.newrelic.com -Port 443
