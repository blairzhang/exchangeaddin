Write-Warning "Please make sure your have well configured the variables in this script."

$EXDIR="C:\Program Files\Microsoft\Exchange Server"
$AgentName="SCOPIA Meeting Routing Agent"
$AgentInstallDir="$EXDIR\TransportRoles\Agents\RoutingAgents\RvScopiaMeetingAddIn"


try{

    Write-Output "Creating directories"
    New-Item -Type Directory -path $AgentInstallDir\Log  -ErrorAction SilentlyContinue

    Write-Output "Copying files"
    Copy-Item RvScopiaMeetingAddIn.dll $AgentInstallDir -force
    Copy-Item RvScopiaMeetingAddIn.pdb $AgentInstallDir -force
    Copy-Item settings.properties $AgentInstallDir -force
    Copy-Item messages.properties $AgentInstallDir -force


    Write-Output "Registering agent"
    Install-TransportAgent -Name $AgentName -AssemblyPath $AgentInstallDir\RvScopiaMeetingAddIn.dll -TransportAgentFactory Radvision.Scopia.ExchangeMeetingAddIn.RvScopiaMeetingFactory

    Write-Output "Enabling agent"
    Enable-TransportAgent -Identity $AgentName

    Get-TransportAgent -Identity $AgentName

    Write-Output "Finished installation."

}catch [System.IO.IOException]
{
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red

    tasklist /m RvScopiaMeetingAddIn.dll

    write-host "Install failed, please try again after the exception above is resolved" -ForegroundColor Red
}
catch
{
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
    write-host "Install failed, please try again after the exception avove is resolved" -ForegroundColor Red
}

