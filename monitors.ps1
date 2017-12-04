$os = "Windows (7|10|XP)"
$windowscomputers = get-adcomputer -filter * -prop * | where-object operatingsystem -match "Windows (7|10)" | select name, lastlogondate, operatingsystem | select -first 50

function global:convertto-char 
(	
	$array
)
{
	$output = ""
	foreach ($char in $array)
	{	
		$output += [char]$char -join ""
	}
	return $output
}

workflow collectinfo
{   
    param
	(
    	[parameter(mandatory=$true, position=0)]
    	[alias("comp")]$computers

	)

    set-psworkflowdata -psallowredirection $true

	$today = get-date
	$validdate = $today.AddMonths(-3)

	$username = "james@pncc.govt.nz"
	$password = get-content d:\james\aenigma | convertto-securestring
	
	$cred = [pscredential]::new($username, $password) 
	write-output $cred | select *
	
	foreach -parallel ($computer in $computers)
	{

		#new-item -path "d:/james/temp/$computer.name" -type File
		if (($computer.lastlogondate -gt $validdate) -and (test-connection -count 1 -comp $computer.name -ea silentlycontinue))
		{   
			try
			{
				$computername = $computer.name
				$loggedinuser = inlinescript {invoke-command -computername $using:computername -credential $using:cred -scriptblock {(get-ciminstance -class cim_computersystem -ea silentlycontinue).username}}
				write-output "Computer $computer.name - User $loggedinuser"
				#write-output "after log user"
				$monitors = inlinescript {invoke-command -comp $using:computer.name -credential $using:cred -scriptblock {get-ciminstance -class win32_pnpentity | where-object service -eq monitor}}
				
				# need to add client network test before cim call and db connect code
				$moncount = 0
				
				foreach ($monitor in $monitors)
				{
					$moncount++
					$monitorid = $monitor.PNPDeviceID + "_0"
					$monitorinfo = inlinescript {invoke-command -computername $using:computername -credential $using:cred -scriptblock {get-ciminstance -class wmimonitorid -namespace root\wmi | Where-Object instancename -eq $($args[0])} -ArgumentList $using:monitorid}
                
					$name = $monitor.name
					#$manufacturer = ($minfo.ManufacturerName -notmatch 0 | foreach-object {[char]$_}) -join ""
					$makemodelchar = global:convertto-char($monitorinfo.UserFriendlyName)

					$makemodel = $makemodelchar.split(" ")
					$make = $makemodel[0]
					$model = $makemodel[1]
					$serial = global:convertto-char($monitorinfo.SerialNumberID)
					$manufacturedate = "Week " + $monitorinfo.WeekOfManufacture + " " + $monitorinfo.YearOfManufacture
					
					write-output "Monitor: $moncount"
				    write-output "Name: $name"
					#write-output "Manufacturer: $manufacturer"
					write-output "Make: $make"
					write-output "Model: $model"
					write-output "Serial: $serial"
					write-output "Manufacture Date: $manufacturedate"
					#test run on domain
				}
			}
			catch [exception] { write-output $computer.name "refused cim call from domain admin!" }
			$monitors = $null
		}
	} 
}

