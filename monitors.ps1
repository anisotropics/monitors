$os = "Windows (7|10|XP)"
$windowscomputers = get-adcomputer -filter * -prop * | where-object operatingsystem -match "Windows (7|10)" | select name, lastlogondate, operatingsystem | select -first 10

$today = get-date
$validdate = $today.AddMonths(-3)

$cred = get-credential -username "james" -message "Password?"

function convertto-char
(	
	$array
)
{
	$output = ""
	foreach ($char in $array)
	{	
		$output += [char]$char -join ""
	}
	return $Output
}


foreach ($computer in $windowscomputers)
{
	if (($computer.lastlogondate -gt $validdate) -and (test-connection -count 1 -comp $computer.name -ea silentlycontinue))
    {   
        try
        {
			$loggedinuser = invoke-command -comp $computer.name -cred $cred -script {(get-ciminstance -class cim_computersystem -ea silentlycontinue).username}
			write-host "Computer $computer.name - User $loggedinuser" -fore green
			$monitors = invoke-command -comp $computer.name -cred $cred -script {get-ciminstance -class win32_pnpentity -comp $computer.name | where-object service -eq monitor}
			
			# need to add client network test before cim call and db connect code
			$moncount = 0
			
			foreach ($monitor in $monitors)
			{
				$moncount++
				$global:mid = $monitor.PNPDeviceID + "_0"
				$minfo = invoke-command -comp $computer.name -cred $cred -script {get-ciminstance -class wmimonitorid -namespace root\wmi | Where-Object instancename -eq $($args[0])} -ArgumentList $mid
			
				$name = $monitor.name
				#$manufacturer = ($minfo.ManufacturerName -notmatch 0 | foreach-object {[char]$_}) -join ""
				$makemodel = convertto-char($minfo.UserFriendlyName)
				$serial = ($minfo.SerialNumberID -notmatch 0 | foreach-object {[char]$_}) -join ""
				$manufacturedate = "Week " + $minfo.WeekOfManufacture + " " + $minfo.YearOfManufacture
				
				write-host "Monitor: $moncount"
				write-host "Name: $name"
				#write-host "Manufacturer: $manufacturer"
				write-host "Make/Model: "$ufn
				write-host "Serial: $serial"
				write-host "Manufacture Date: $manufacturedate `n"
				#test run on domain
			}
        }
        catch [exception] { write-host $computer.name "refused cim call from domain admin!" }
	}
} 


