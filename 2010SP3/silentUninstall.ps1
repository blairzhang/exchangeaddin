Write-Warning "Please make sure your have well configured the variables in this script."


$EXDIR="C:\Program Files\Microsoft\Exchange Server"
$AgentName="SCOPIA Meeting Routing Agent"
$AgentInstallDir="$EXDIR\TransportRoles\Agents\RoutingAgents\RvScopiaMeetingAddIn"

$ErrorActionPreference = "Stop"
try{

	Get-TransportAgent
	Write-Output "Disabling Agent..."
	Disable-TransportAgent -Identity $AgentName -Confirm:$false

	Write-Output "Uninstalling Agent.."
	Uninstall-TransportAgent -Identity $AgentName -Confirm:$false
    

        Write-Output "Restarting MSExchangeTransport..."
        Restart-Service MSExchangeTransport
    
        Write-Output "Reset IIS service..."
        IISRESET

	Write-Output "Deleteing Files and Folders..."
	Remove-Item $AgentInstallDir\* -Recurse -ErrorAction SilentlyContinue
	Remove-Item $AgentInstallDir -Recurse -ErrorAction SilentlyContinue

	Get-TransportAgent
	Write-Output "Transport Agent successfully uninstalled."
}catch [System.IO.IOException]
{
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red

    tasklist /m RvScopiaMeetingAddIn.dll

    write-host "Uninstall failed, please try again after the exception avove is resolved" -ForegroundColor Red

}
catch
{
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
    write-host "Install failed, please try again after the exception avove is resolved" -ForegroundColor Red

}






