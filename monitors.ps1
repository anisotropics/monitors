$windowscomputers = get-adcomputer -filter * -prop * | where operatingsystem -match "Windows (7|10|XP)" | select name, lastlogondate, operatingsystem | select -first 10

$today = get-date


foreach ($computer in $windowscomputers)
{
	if (test-connection -count 1 -comp $computer.name -ea silentlycontinue)
    {   
        try
        {
			$loggedinuser = (get-ciminstance -class cim_computersystem -comp $computer.name -ea silentlycontinue).username
			write-host "Computer $computer.name - User $loggedinuser" -fore green
			$monitors = get-ciminstance -class win32_pnpentity -comp $computer.name | where-object service -eq monitor
			
			# need to add client network test before cim call and db connect code
			$moncount = 0
			
			foreach ($monitor in $monitors)
			{
				$moncount++
				$mid = $monitor.PNPDeviceID + "_0"
				$minfo = get-ciminstance -class wmimonitorid -namespace root\wmi | Where-Object instancename -eq $mid
			
				$name = $monitor.name
				$manufacturer = ($minfo.ManufacturerName -notmatch 0 | foreach-object {[char]$_}) -join ""
				$serial = ($minfo.SerialNumberID -notmatch 0 | foreach-object {[char]$_}) -join ""
				$manufacturedate = "Week " + $minfo.WeekOfManufacture + " " + $minfo.YearOfManufacture
				
				write-host "Monitor: $moncount"
				write-host "Name: $name"
				write-host "Manufacturer: $manufacturer"
				write-host "Serial: $serial"
				write-host "Manufacture Date: $manufacturedate `n"
				#test run on domain
			}
        }
        catch [exception] { write-host $pc "refused wmi call from domain admin!" }
	}
} 


