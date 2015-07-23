$EXDIR="C:\Program Files\Microsoft\Exchange Server"
$AgentName="Radvision SCOPIA Exchange Meeting AddIn"
$AgentInstallDir="$EXDIR\TransportRoles\Agents\RoutingAgents\RvScopiaMeetingAddIn"

Net Stop MSExchangeTransport

Write-Output "Creating directories"
New-Item -Type Directory -path $AgentInstallDir\Log  -ErrorAction SilentlyContinue

Write-Output "Copying files"
Copy-Item bin\Debug\RvScopiaMeetingAddIn.dll $AgentInstallDir -force
Copy-Item bin\Debug\RvScopiaMeetingAddIn.pdb $AgentInstallDir -force
Copy-Item settings.properties $AgentInstallDir -force
Copy-Item messages.properties $AgentInstallDir -force

Write-Output "Registering agent"
Install-TransportAgent -Name $AgentName -AssemblyPath $AgentInstallDir\RvScopiaMeetingAddIn.dll -TransportAgentFactory Radvision.Scopia.ExchangeMeetingAddIn.RvScopiaMeetingFactory

Write-Output "Enabling agent"
Enable-TransportAgent -Identity $AgentName
Get-TransportAgent -Identity $AgentName

Net Start MSExchangeTransport

Write-Output "Install Complete. Please exit the Exchange Management Shell."