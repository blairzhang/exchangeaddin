Write-Output "Restarting MSExchangeTransport..."
Restart-Service MSExchangeTransport
    
Write-Output "Reset IIS service..."
IISRESET

Write-Output "The current transport agents..."
Get-TransportAgent