# Get titles from the anchor links in teasers
$titles = $ParsedHTMLResponse.QuerySelectorAll('a[href*="/vrtnws/nl/"] .teaser-text') | 
          ForEach-Object { $_.TextContent.Trim() } |
          Where-Object { $_ -ne "" }


# Nice formatted output
Write-Host "`n📰 VRT NWS Headlines`n" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Yellow

for ($i = 0; $i -lt $titles.Count; $i++) {
    Write-Host "$($i + 1). " -ForegroundColor Green -NoNewline
    Write-Host $titles[$i] -ForegroundColor White
    if (($i + 1) -lt $titles.Count) {
        Write-Host ("-" * 60) -ForegroundColor DarkGray
    }
}


Write-Host "`nTotal articles found: $($titles.Count)" -ForegroundColor Cyan

