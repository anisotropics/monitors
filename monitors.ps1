#$os = "Windows (7|10|XP)"
$windowscomputers = get-adcomputer -filter * -prop * | where-object operatingsystem -match "Windows (7|10)" | select-object name, lastlogondate, operatingsystem | select-object -first 20

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

class collectedcomputer
{
	[guid]$guid
	[string]$name
	[string]$dateinstalled
	[string]$lastloggedonuser
	[string]$lastlogonts

	collectedcomputer($guid, $name, $dateinstalled, $lastloggedonuser, $lastlogonts)
	{
		$this.guid = $guid
		$this.name = $name
		$this.dateinstalled = $dateinstalled
		$this.lastloggedonuser = $lastloggedonuser
		$this.lastlogonts = $lastlogonts
	}
}

class collectedmonitor
{
	[string]$make
	[string]$model
	[string]$serialnumber
	[string]$manufactureweek
	[string]$manufactureyear

	collectedmonitor($make, $model, $serialnumber, $manufactureweek, $manufactureyear)
	{
		$this.make = $make
		$this.model = $model
		$this.serialnumber = $serialnumber
		$this.manufactureweek = $manufactureweek
		$this.manufactureyear = $manufactureyear
	}
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

	<#------------------
		set up database insert\update objects  
	#>
	function global:mailboxtodb
	{
		foreach ($mailbox in get-mailbox)
		{
			$mailstats = get-mailboxstatistics -Identity $mailbox.userprincipalname
			
			$date = (get-date -Format "yyyy-MM-dd HH:mm:ss").tostring()
			$guid = $mailbox.exchangeguid
			$upn = $mailbox.UserPrincipalName
			$email = $mailbox.WindowsEmailAddress
			$size = $mailstats.TotalItemSize
			$itemcount = $mailstats.ItemCount
			$lastlogon = $mailstats.LastLoggedOnUserAccount
			$name = $mailbox.displayname
			$class = $mailstats.ObjectClass
			$lastlogontimestamp = (get-date $mailstats.LastLogonTime -format "yyyy-MM-dd HH:mm:ss").tostring()
	
			write-host $date
			write-host $lastlogontimestamp
	
			$mailinfo = new-object -type mailboxinfo $guid, $name, $email, $size, $date, $lastlogon, $itemcount, $upn, $class, $lastlogontimestamp
			write-objecttosql -InputObject $mailinfo -Server sqldev3.pncc.govt.nz -Database exchangeinfo -TableName mailboxstats
		}
	
		return $mailinfo
	}

	#-------------------

	
	foreach -parallel ($computer in $computers)
	{

		#new-item -path "d:/james/temp/$computer.name" -type File
		if (($computer.lastlogondate -gt $validdate) -and (test-connection -count 1 -comp $computer.name -ea silentlycontinue))
		{   
			$computername = $computer.name

			try
			{
				$loggedinuser = inlinescript {invoke-command -computername $using:computername -credential $using:cred -scriptblock {(get-ciminstance -class cim_computersystem -ea silentlycontinue).username}}
				write-output "Computer $computer.name - User $loggedinuser"
				#write-output "after log user"
				$monitors = inlinescript {invoke-command -comp $using:computer.name -credential $using:cred -scriptblock {get-ciminstance -class win32_pnpentity | where-object service -eq monitor}}
				$computerobject = new-object -type collectedcomputer $guid, $name, $email, $size, $date, $lastlogon, $itemcount, $upn, $class, $lastlogontimestamp

			}
			catch [exception] { write-output $computer.name "that didn't work!" $error }

			try
			{
				$moncount = 0
		
				foreach ($monitor in $monitors)
				{
					$moncount++
					$monitorid = $monitor.PNPDeviceID + "_0"
					$monitorinfo = inlinescript {invoke-command -computername $using:computername -credential $using:cred -scriptblock {get-ciminstance -class wmimonitorid -namespace root\wmi | where-object instancename -eq $($args[0])} -argumentList $using:monitorid}
                
					$name = $monitor.name
					#$manufacturer = ($minfo.ManufacturerName -notmatch 0 | foreach-object {[char]$_}) -join ""
					$makemodelchar = global:convertto-char($monitorinfo.UserFriendlyName)

					$makemodel = $makemodelchar.split(" ")
					$make = $makemodel[0]
					$model = $makemodel[1]
					$serial = global:convertto-char($monitorinfo.SerialNumberID)
					$manufactureweek = $monitorinfo.WeekOfManufacture
					$manufactureyear = $monitorinfo.YearOfManufacture
					
					write-output "Monitor: $moncount"
				    write-output "Name: $name"
					write-output "Make: $make"
					write-output "Model: $model"
					write-output "Serial: $serial"
					write-output "Manufacture Date: $manufacturedate"
					#test run on domain
				}
			}
			catch [exception] { write-output $computer.name "that didn't work!" $error }
			$monitors = $null
		}
	} 
}

collectinfo $windowscomputers