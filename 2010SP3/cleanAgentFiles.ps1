Write-Warning "Please make sure your have well configured the variables in this script."


$EXDIR="C:\Program Files\Microsoft\Exchange Server"
$AgentName="SCOPIA Meeting Routing Agent"
$AgentInstallDir="$EXDIR\TransportRoles\Agents\RoutingAgents\RvScopiaMeetingAddIn"
	
Write-Output "Deleteing Files and Folders..."
Remove-Item $AgentInstallDir\* -Recurse -ErrorAction SilentlyContinue
Remove-Item $AgentInstallDir -Recurse -ErrorAction SilentlyContinue
Write-Output "Finished deleting."







