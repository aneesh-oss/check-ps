New-NetFirewallRule -DisplayName "Allow New Relic Download" `
    -Direction Outbound -Action Allow -RemoteAddress "download.newrelic.com"

Resolve-DnsName download.newrelic.com

Test-NetConnection download.newrelic.com -Port 443


& "C:\Program Files\New Relic\newrelic-infra\newrelic-integrations\nri-flex.exe" --verbose --pretty --config_path "C:\Program Files\New Relic\newrelic-infra\integrations.d\exchange-config.yml"
