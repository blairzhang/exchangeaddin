$EXDIR="C:\Program Files\Microsoft\Exchange Server"
$AgentName="Radvision SCOPIA Exchange Meeting AddIn"
$AgentInstallDir="$EXDIR\TransportRoles\Agents\RoutingAgents\RvScopiaMeetingAddIn"


Net Stop MSExchangeTransport


Write-Output "Disabling Agent..."
Disable-TransportAgent -Identity $AgentName -Confirm:$false

Write-Output "Uninstalling Agent.."
Uninstall-TransportAgent -Identity $AgentName -Confirm:$false

Write-Output "Deleteing Files and Folders..."
Remove-Item $AgentInstallDir\* -Recurse -ErrorAction SilentlyContinue
Remove-Item $AgentInstallDir -Recurse -ErrorAction SilentlyContinue

Net Start MsExchangeTransport

Write-Output "Uninstall Complete."
