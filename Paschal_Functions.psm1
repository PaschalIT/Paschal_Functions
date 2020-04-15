# Paschal Powershell Functions. 
# To be deployed on all Paschal Computers and Servers. 
# To reference this function file, run the following command. 
# Import-Module Paschal_Functions

<#		Find Replace in external file
$file = '<File Path>'
$find = <Query>			Regex for anything after query before CR '[^"]*'
$replace = <New>
if ((Get-Content $file) -like $find)
{
	(Get-Content $file) -replace $find, $replace | Set-Content $file
}
else
{
	Add-Content $file $replace
}
#>

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

function PSleep([int]$seconds, [string]$action, [string]$description) {
	# See below for use. 
	# EXAMPLE - psleep 15 Syncing "Synchronizing with Office365"
	$doneDT = (Get-Date).AddSeconds($seconds)
	while ($doneDT -gt (Get-Date)) {
		$secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
		$percent = ($seconds - $secondsLeft) / $seconds * 100
		Write-Progress -Activity "$description" -Status "$action..." -SecondsRemaining $secondsLeft -PercentComplete $percent
		[System.Threading.Thread]::Sleep(100)
	}
	Write-Progress -Activity "$action" -Status "$action..." -SecondsRemaining 0 -Completed
}

function PPS([string]$computer) {
	if ($computer) {
	} else {
		$computer = (Read-Host -Prompt `n'Computer to connect to?')
	}
	if ($env:username -match '.+?\-adm$' -or $env:username -match 'srv') {
		Enter-PSSession -ComputerName $computer -Credential (Get-Credential $env:username)
	} else {
		Enter-PSSession -ComputerName $computer -Credential (Get-Credential "$($env:username)-adm")
	}
}

function ADSync {
	# Force AD Sync to Office 365
	Write-Host ""
	Read-Host -Prompt "Press Enter to sync with O365"
	Write-Host "Syncing to Office 365, Please Wait....."
	powershell.exe "Invoke-Command -ComputerName WDC01V -ScriptBlock  {import-module 'C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync\ADSync.psd1' ; Start-ADSyncSyncCycle -PolicyType Initial}"
	Write-Host "Waiting For Sync to Complete"
	psleep 60 Syncing "Synchronizing with Office365"
	Write-Host ""
	Write-Host -ForegroundColor Green "Done!"
	Write-Host ""
}

function PListSelect # Must pass in array of Strings!  Function returns an array with the selected strings.  Format - PListSelect [String Array] [Limit]
{
	<#
        .Synopsis
        Select one or more items from a list of options.

		.Description
		PListSelect takes an input array, displays all of the options, and lets the user select one or more.  The selected items are then placed in a new array which outputs as the return.

        .Example
        PListSelect $mylist -limit 3

        .Parameter list
        An array of strings.

        .Parameter limit
        An int value.

        .Parameter customprompt
        A string.
	#>
	
	[Alias("pls")]
	param ([parameter(Mandatory = $true)]
		[string[]]$list = @(),
		[int]$limit = 0,
		[string]$prompt = $null,
		[boolean]$pclear = $false)
	
	$select = @(); $ret = @(); $offset = 0
	
	if ($list.Length -eq 0) {
		if (!$pclear) {
			Clear-Host
		}
		write-host -ForegroundColor Red "No input array provided.  Array or multiple values required.  Press enter to continue."
		read-host
		break
	}
	
	foreach ($i in $list) {
		$select += $false
	}
	
	do {
		if (!$pclear) {
			Clear-Host
		}
		$count = 1 # Trash variable to track numbering
		
		if ($prompt) {
			write-host -ForegroundColor Cyan $prompt
			write-host ""
		}
		
		for ($i = 0; $i -lt 25; $i++) # Write out options in columns of 15
{
			if (!$list[($i + $offset)]) {
				break
			}
			
			write-host -NoNewline ($i + $offset + 1); write-host -NoNewline ")`t"
			if ($select[($i + $offset)]) {
				write-host -ForegroundColor Green (($list[$i + $offset].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
			} else {
				write-host (($list[$i + $offset].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
			}
			if ($list[($i + $offset)].Length -gt 20) {
				write-host "..." -NoNewLine # If more than 20 characters, adds ellipses
			}
			
			if ($list[($i + $offset + 25)]) # Checks to see if next column is needed; prints items 26-50 if they exist
{
				write-host -NoNewLine "`t"
				if ($list[($i + $offset)].Length -lt 8) # These are to properly align everything in columns
{
					write-host -NoNewline "`t"
				}
				if ($list[($i + $offset)].Length -lt 16) {
					write-host -NoNewLine "`t"
				}
				write-host -NoNewline ($i + $offset + 26); write-host -NoNewline ")`t"
				if ($select[($i + $offset + 25)]) {
					write-host -ForegroundColor Green (($list[($i + $offset + 25)].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
				} else {
					write-host (($list[($i + $offset + 25)].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
				}
				if ($list[($i + $offset + 25)].Length -gt 20) {
					write-host "..." -NoNewLine # If more than 20 characters, adds ellipses
				}
			}
			
			if ($list[($i + $offset + 50)]) # Checks to see if next column is needed; prints items 51-75 if they exist
{
				write-host -NoNewLine "`t"
				if ($list[($i + $offset + 25)].Length -lt 8) # These are to properly align everything in columns
{
					write-host -NoNewLine "`t"
				}
				if ($list[($i + $offset + 25)].Length -lt 16) {
					write-host -NoNewLine "`t"
				}
				write-host -NoNewLine ($i + $offset + 51); write-host -NoNewLine ")`t"
				if ($select[($i + $offset + 50)]) {
					write-host -ForegroundColor Green (($list[($i + $offset + 50)].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
				} else {
					write-host (($list[($i + $offset + 50)].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
				}
				if ($list[($i + $offset + 50)].Length -gt 20) {
					write-host "..." -NoNewLine # If more than 20 characters, adds ellipses
				}
			}
			
			write-host ""
		}
		
		#foreach ($i+$offset in $list)
		#{
		#write-host -NoNewLine "$count)`t"
		#if ($select[($count-1)])
		#{
		#write-host -f Green $i+$offset  # Write green if selected
		#}
		#else
		#{
		#write-host $i+$offset  # Write normal if not selected
		#}
		#$count += 1
		#}
		
		write-host -ForegroundColor Cyan "`r`nPlease select an item.  Selecting a highlighted item will deselect it.  Use 'A' to select all, or 'D' to deselect all.  Enter 'Y' when finished." -NoNewLine
		if ($limit) {
			write-host -ForegroundColor Red "  Limit of $limit selections." -NoNewLine
		}
		if ($list.Length -gt 75) {
			write-host ""
			if (($offset + 75) -lt $list.Length) {
				write-host -ForegroundColor Cyan "Use 'N' to display the next page.  " -NoNewLine
			}
			if (($offset - 75) -ge 0) {
				write-host -ForegroundColor Cyan "Use 'P' to display the previous page." -NoNewLine
			}
		}
		$x = read-host
		
		try {
			$x = [int]$x # Attempts to parse input to int.  Does nothing if input is not numerical.
		} catch [System.Management.Automation.PSInvalidCastException] {
		} # Prevent parse error from displaying.  It does not effect anything in the code.
		
		if ($x -match "^\d+$" -and $x -gt 0 -and $x -le $list.Length) {
			$x -= 1 # Make variable match array values
			
			if ($select[$x]) # If true set false, and vice versa
{
				$select[$x] = $false
			} else {
				if (($select | Where-Object -FilterScript {
							$_ -eq $true
						}).Count -lt $limit -or $limit -eq 0) {
					$select[$x] = $true
				} else {
					write-host -ForegroundColor Red "`r`nCan't make selection as it exceeds the set limit of $limit items.  Please press enter and deselect one before choosing another."
					$dump = read-host
				}
			}
		} elseif ($x -eq 'd') # Deselect all
{
			for ($i = 0; $i -lt $select.Length; $i++) {
				$select[$i] = $false
			}
		} elseif ($x -eq 'a') # Select all
{
			if ($limit -ge $list.Length -or $limit -eq 0) {
				for ($i = 0; $i -lt $select.Length; $i++) {
					$select[$i] = $true
				}
			} else {
				write-host -ForegroundColor Red "`r`nCan't select all values as it exceeds the set limit of $limit items.  Please press enter and select individual values."
				$dump = read-host
			}
		} elseif ($x -eq 'n') {
			if (($offset + 75) -lt $list.Length) {
				$offset += 75
			}
		} elseif ($x -eq 'p') {
			if (($offset - 75) -ge 0) {
				$offset -= 75
			}
		} elseif ($x -ne 'y') {
			write-host -ForegroundColor Red "`r`nInput outside of available selection range.  Please press enter and try again."
			$dump = read-host
		}
	} while ($x -ne 'y') # Repeat until user keys Y
	
	for ($i = 0; $i -lt $select.Length; $i++) # Fill return array with selected strings
{
		if ($select[$i] -and $list[$i] -and $i -lt $list.Length) {
			$ret += $list[$i]
		}
	}
	
	if ($ret.Length -eq 0) {
		return $null
	} else {
		return $ret # Output return array
	}
}

function PSelect {
	<#
	.Synopsis
	Select one item from a list of options.
	
	.Description
	Pass in a list of items and select one, which is output from the function.
	
	.Example
	$outputvar = PSelect $mylist
	#>
	
	[Alias("psel")]
	param ([parameter(Mandatory = $true)]
		[string[]]$list = @(),
		[string]$prompt = $null,
		[boolean]$pclear = $false)
	
	if ($list.Length -le 0) # Check that array exists and isn't empty
{
		if (!$pclear) {
			Clear-Host
		}
		write-host -ForegroundColor Red "Input list is empty or missing.  Please check your function call."
		$dump = read-host
		break
	}
	
	$check = $false
	
	do {
		if (!$pclear) {
			Clear-Host
		}
		$count = 1
		
		if ($prompt) # Print custom prompt if it exists, else print nothing
{
			write-host ""
			write-host -ForegroundColor Cyan $prompt
			write-host ""
		}
		
		foreach ($i in $list) # Print list of options with numbering
{
			write-host -NoNewLine "$count)`t"
			write-host $i
			
			$count += 1
		}
		
		write-host "`r`n99)`tBack/Exit/Cancel"
		
		write-host -ForegroundColor Cyan "`r`nPlease select an item.  " -NoNewLine
		$x = read-host
		
		try {
			$x = [int]$x
		} catch {
		}
		
		if ($x -match "^\d+$" -and (($x -gt 0 -and $x -le $list.Length) -or $x -eq 99)) # Make sure user input is number and within range
{
			return ($x - 1) # Returns user input minus 1 so it matches proper array value
		} else # Error if invalid input
{
			write-host -ForegroundColor Red "`r`nInvalid selection.  Please press enter and try again."
			$dump = read-host
		}
	} while (!$check) # Repeat until correct input
}

function PDebug {
	
	[Alias("pd")]
	param ()
	
	Start-Sleep -Seconds 2
	
	if ([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftAlt)) {
		if ([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::D2) -or [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::NumPad2)) {
			Set-PSDebug -Trace 2
		} else {
			Set-PSDebug -Trace 1
		}
		return $true
	} elseif ([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftCtrl)) {
		return $true
	} else {
		Set-PSDebug -Off
		return $false
	}
}

function PInput {
	param ([parameter(Mandatory = $true)]
		[string]$prompt = $null,
		[int]$req = -1,
		[int]$allownull = 0)
	
	if ($req -eq -1 -and !$allownull) {
		$allownull = 1
	}
	
	do {
		write-host $prompt
		$var = read-host
		
		if (($allownull -and $var.Length -eq 0) -or ($var.Length -eq $req -and $req -gt 0) -or $req -eq -1 -or ($req -eq 0 -and $var.Length -gt 0)) # if (null allowed and input null) (requirement exists and input matches requirement) (no requirement, so anything goes) (requirement is 0 so input not allowed, input exists and is not null)
{
			return $var # If required conditions are met, return the input
		} elseif (!$var) {
			write-host ""
			if (!$Script:pclear) {
				Clear-Host
			}
			write-host -ForegroundColor Red "Input required.  Please try again.`r`n"
		} else {
			write-host ""
			if (!$Script:pclear) {
				Clear-Host
			}
			write-host -ForegroundColor Red "Required length not met.  Please try again.`r`n"
		}
	} while (2 -lt 3) # Repeat indefinitely until return conditions are met
}

function PTitle {
	param ([string]$title = "Paschal IT",
		[string]$version = "0.0")
	
	$Host.UI.RawUI.WindowTitle = "$title v$version"
}

filter Get-InstalledSoftware ## Courtesy of Chris Dent, Powershell Guru
{
    <#
    .SYNOPSIS
        Get all installed from the Uninstall keys in the registry.
    .DESCRIPTION
        Read a list of installed software from each Uninstall key.

        This function provides an alternative to using Win32_Product.
    .EXAMPLE
        Get-InstalledSoftware

        Get the list of installed applications from the local computer.
    .EXAMPLE
        Get-InstalledSoftware -IncludeLoadedUserHives

        Get the list of installed applications from the local computer, including each loaded user hive.
    .EXAMPLE
        Get-InstalledSoftware -ComputerName None -DebugConnection

        Display all error messages thrown when attempting to audit the specified computer.
    .EXAMPLE
        Get-InstalledSoftware -IncludeBlankNames

        Display all results, including those with very limited information.
    #>
	
	[CmdletBinding()]
	[OutputType([PSObject])]
	param (
		# The computer to execute against. By default, Get-InstalledSoftware reads registry keys on the local computer.
		[Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[String]$ComputerName = $env:COMPUTERNAME,
		# Attempt to start the remote registry service if it is not already running. This parameter will only take effect if the service is not disabled.

		[Switch]$StartRemoteRegistry,
		# Some software packages, such as DropBox install into a users profile rather than into shared areas. Get-InstalledSoftware can increase the search to include each loaded user hive.

		#

		# If a registry hive is not loaded it cannot be searched, this is a limitation of this search style.

		[Switch]$IncludeLoadedUserHives,
		# By default Get-InstalledSoftware will suppress the display of entries with minimal information. If no DisplayName is set it will be hidden from view. This behaviour may be changed using this parameter.

		[Switch]$IncludeBlankNames
	)
	
	$keys = 'Software\Microsoft\Windows\CurrentVersion\Uninstall',
	'Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
	
	# If the remote registry service is stopped before this script runs it will be stopped again afterwards.
	if ($StartRemoteRegistry) {
		$shouldStop = $false
		$service = Get-Service RemoteRegistry -ComputerName $ComputerName
		
		if ($service.Status -eq 'Stopped' -and $service.StartType -ne 'Disabled') {
			$shouldStop = $true
			$service | Start-Service
		}
	}
	
	$baseKeys = [System.Collections.Generic.List[Microsoft.Win32.RegistryKey]]::new()
	
	$baseKeys.Add([Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName, 'Registry64'))
	if ($IncludeLoadedUserHives) {
		try {
			$baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users', $ComputerName, 'Registry64')
			foreach ($name in $baseKey.GetSubKeyNames()) {
				if (-not $name.EndsWith('_Classes')) {
					Write-Debug ('Opening {0}' -f $name)
					
					try {
						$baseKeys.Add($baseKey.OpenSubKey($name, $false))
					} catch {
						$errorRecord = [System.Management.Automation.ErrorRecord]::new(
							$_.Exception.GetType()::new(
								('Unable to access sub key {0} ({1})' -f $name, $_.Exception.InnerException.Message.Trim()),
								$_.Exception
							),
							'SubkeyAccessError',
							'InvalidOperation',
							$name
						)
						Write-Error -ErrorRecord $errorRecord
					}
				}
			}
		} catch [Exception] {
			Write-Error -ErrorRecord $_
		}
	}
	
	foreach ($baseKey in $baseKeys) {
		Write-Verbose ('Reading {0}' -f $baseKey.Name)
		
		if ($basekey.Name -eq 'HKEY_LOCAL_MACHINE') {
			$username = 'LocalMachine'
		} else {
			# Attempt to resolve a SID
			try {
				[System.Security.Principal.SecurityIdentifier]$sid = Split-Path $baseKey.Name -Leaf
				$username = $sid.Translate([System.Security.Principal.NTAccount]).Value
			} catch {
				$username = Split-Path $baseKey.Name -Leaf
			}
		}
		
		foreach ($key in $keys) {
			try {
				$uninstallKey = $baseKey.OpenSubKey($key, $false)
				
				if ($uninstallKey) {
					$is64Bit = $true
					if ($key -match 'Wow6432Node') {
						$is64Bit = $false
					}
					
					foreach ($name in $uninstallKey.GetSubKeyNames()) {
						$packageKey = $uninstallKey.OpenSubKey($name)
						
						$installDate = Get-Date
						$dateString = $packageKey.GetValue('InstallDate')
						if (-not $dateString -or -not [DateTime]::TryParseExact($dateString, 'yyyyMMdd', (Get-Culture), 'None', [Ref]$installDate)) {
							$installDate = $null
						}
						
						[PSCustomObject]@{
							Name	    = $name
							DisplayName = $packageKey.GetValue('DisplayName')
							DisplayVersion = $packageKey.GetValue('DisplayVersion')
							InstallDate = $installDate
							InstallLocation = $packageKey.GetValue('InstallLocation')
							HelpLink    = $packageKey.GetValue('HelpLink')
							Publisher   = $packageKey.GetValue('Publisher')
							UninstallString = $packageKey.GetValue('UninstallString')
							URLInfoAbout = $packageKey.GetValue('URLInfoAbout')
							Is64Bit	    = $is64Bit
							Hive	    = $baseKey.Name
							Path	    = Join-Path $key $name
							Username    = $username
							ComputerName = $ComputerName
						}
					}
				}
			} catch {
				Write-Error -ErrorRecord $_
			}
		}
	}
	
	# Stop the remote registry service if required  
	if ($StartRemoteRegistry -and $shouldStop) {
		$service | Stop-Service
	}
}

function Enable-PaschalEXCContacts {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $false, ValueFromPipeline = $true)]
		[ValidatePattern("@gopaschal\.com|@paschalcorp\.com")]
		[string[]]$TargetMailbox,
		[Parameter(Mandatory = $false)]
		[PSCredential]$Credentials = (Get-Credential srv -Message "Enter Exchange Server credentials.")
	)
	
	$mailboxtotal = $TargetMailbox.Count; $mailboxcount = 1;
	
	if (-not (Get-Module PSFramework -ListAvailable -ErrorAction SilentlyContinue)) {
		Install-Module PSFramework -Confirm:$false
	}
	Import-Module PSFramework
	
	$userlist = Get-ADUser -Filter * -SearchBase "OU=Springdale, DC=US, DC=PaschalCorp, DC=com" -Properties Mobile, TelephoneNumber, Mail, GivenName, Surname, Department, Title, ProxyAddresses | Where-Object {
		$_.Mobile -or $_.TelephoneNumber -or $_.SamAccountName -eq "_temp"
	}
	
	$contacthash = $userlist |
	Where-Object {
		$_.DistinguishedName -match "OU=Users" -or $_.Name -eq "CSR On Call"
	} |
	Select-PSFObject @(
		"GivenName as FirstName"
		"Surname as LastName"
		"Name as DisplayName"
		"Mobile as MobilePhone"
		"TelephoneNumber as BusinessPhone"
		"Mail as EmailAddress"
		"SID as Mileage"
		"Title"
		"Department"
	) |
	ConvertTo-PSFHashtable -Include FirstName, LastName, BusinessPhone, MobilePhone, DisplayName, Department, Title, EmailAddress, Mileage
	
	Clear-Host
	Write-Host -ForegroundColor White "# of Users to convert to Contacts - $($contacthash.Count)"
	
	foreach ($mailbox in $TargetMailbox) {
		
		$targetUser = $userlist | Where-Object {
			$_.ProxyAddresses -match $mailbox
		}
		$targetUser = Get-ADUser $targetUser.SamAccountName -Properties DistinguishedName, Mail
		
		Write-Host -ForegroundColor White "$($mailbox) - $mailboxcount of $mailboxtotal"
		Write-Host ""; Write-Host ""
		
		Invoke-Command -ComputerName WEM01V -Credential $exchcred -ScriptBlock {
			
			$success = 0; $SIDfailure = 0; $totalfailure = 0;
			$contacthash = $args[0]; $targetUser = $args[1]; $exchcred = $args[2]; $count = 1;
			
			Write-Host -ForegroundColor White "Deploying $($contacthash.Count) Contacts to $($targetUser.Mail)"
			
			foreach ($contact in $contacthash) {
				Write-Host -ForegroundColor White ($current = "$($contact.DisplayName) ($count of $($contacthash.Count)) -> $($targetUser.Mail)")
				for ($i = 0; $i -lt $current.Length; $i++) {
					Write-Host -ForegroundColor White '-' -NoNewline
				}
				Write-Host ""
				
				try {
					New-EXCContact -MailboxName $targetUser.Mail -MailboxOwnerDistinguishedName $targetUser.DistinguishedName -FirstName $contact.FirstName -LastName $contact.LastName -DisplayName $contact.DisplayName -BusinessPhone $contact.BusinessPhone -MobilePhone $contact.MobilePhone -EmailAddress $contact.EmailAddress -Department $contact.Department -JobTitle $contact.Title -CompanyName "Paschal Air, Plumbing & Electric - AR01-PS" -Credentials $exchcred -useImpersonation
					Write-Host -ForegroundColor Green "Complete!"
					# Start-Sleep -Seconds 1
					Write-Host -ForegroundColor Yellow "Retrieving Contact..."
					$temp = Get-EXCContacts -MailboxName $targetUser.Mail -Credentials $exchcred -useImpersonation | Where-Object {
						$_.DisplayName -eq $contact.DisplayName
					}
					Write-Host -ForegroundColor Yellow "Setting SID as Mileage property..."
					
					try {
						$temp.Mileage = $contact.Mileage
						$temp.Update("AutoResolve")
						Write-Host -ForegroundColor Green "Set SID as Mileage property successfully!"
						$success++
					} catch {
						Write-Host -ForegroundColor Red "Could not set SID as Mileage property..."
						$SIDfailure++
					}
				} catch {
					Write-Host -ForegroundColor Red "Contact creation failed for $($contact.DisplayName) -> $($targetUser.Mail)"
					$totalfailure++
				}
				
				Write-Host ""
				$count++
			}
			Write-Host ""
			
			Write-Host -ForegroundColor White "Process Completed"
			Write-Host -ForegroundColor Green "Contacts Created Successfully- $success"
			Write-Host -ForegroundColor Yellow "Contacts Created with SID Failure - $SIDfailure"
			Write-Host -ForegroundColor Red "Contacts Failed upon Creation - $totalfailure"
		} -ArgumentList $contacthash, $targetUser, $exchcred
		
		$mailboxcount++
	}
	
}

function Update-PaschalEXCContacts {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false, ValueFromPipeline = $true)]
		[ValidatePattern("@gopaschal\.com|@paschalcorp\.com")]
		[string[]]$TargetMailboxes = @((Get-ADUser -Filter * -SearchBase "OU=Users, OU=Springdale, DC=US, DC=PaschalCorp, DC=com" -Properties Mail, Mobile | Where-Object {
					$_.Mobile
				}).Mail + (Get-ADUser -Identity csr-on-call -Properties Mail).Mail),
		[Parameter(Mandatory = $false)]
		[PSCredential]$Credentials = (Get-Credential srv)
	)
	
	$ErrorActionPreference = 'Continue'
	
	# $dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
	# [void][Reflection.Assembly]::LoadFile($dllpath)
	
	$exchcred = $Credentials
	if (-not $exchcred) {
		$exchcred = Get-Credential srv
	}
	$userlist = Get-ADUser -Filter * -SearchBase "OU=Springdale, DC=US, DC=Paschalcorp, DC=com" -Properties Mail, Mobile, TelephoneNumber | Where-Object {
		$_.Mail -and ($_.Mobile -or $_.TelephoneNumber) -and $_.Name -notmatch "ADM"
	}
	$emailtext = $null; $mailboxcount = 1;
	
	Clear-Host
	foreach ($mailbox in $TargetMailboxes) {
		$updatecount = 0; $updatefailed = @(); $contactlist = @();
		
		Write-Host -ForegroundColor White ($text = "Current Target - $mailbox ($mailboxcount of $($TargetMailboxes.Count))")
		for ($i = 0; $i -lt $text.Length; $i++) {
			Write-Host -ForegroundColor White '-' -NoNewline
		}
		Write-Host ""
		
		$contactlist = Get-EXCContacts -MailboxName $mailbox -Credentials $exchcred -useImpersonation | Where-Object {
			$_.CompanyName -match "AR01-PS"
		}
		
		if (-not $contactlist) {
			$emailtext += @"
$mailbox
Could not update Contacts
Unable to locate Contacts on Exchange Mailbox

"@
		} else {
			foreach ($contact in $contactlist) {
				$user = $null;
				if ($contact.Mileage) {
					$user = $userlist | Where-Object {
						$_.SID -eq $contact.Mileage
					}
				} else {
					$user = $userlist | Where-Object {
						$(
							$count = 0
							if ($_.GivenName -eq $contact.GivenName) {
								$count++
							}
							if ($_.Surname -eq $contact.Surname) {
								$count++
							}
							if ($_.Name -eq $contact.DisplayName) {
								$count++
							}
							if ($_.Mobile -eq $contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone]) {
								$count++
							}
							if ($_.Mail -eq $contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address) {
								$count++
							}
							$count
						) -ge 3
					}
				}
				
				if (-not $user) {
					Write-Host -ForegroundColor Red "Contact $($contact.DisplayName) was not found in current list of Active Directory users."
					$updatefailed += "$($contact.DisplayName)(Locate)"
				} else {
					try {
						$itemsupdated = @()
						$contact.Mileage = $user.SID
						$itemsupdated += "SID"
						$contact.DisplayName = $user.Name
						$itemsupdated += "DisplayName"
						if ($user.Mobile) {
							$contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $user.Mobile
							$itemsupdated += "MobilePhone"
						}
						if ($user.TelephoneNumber) {
							$contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $user.TelephoneNumber
							$itemsupdated += "BusinessPhone"
						}
						if ($user.Mail) {
							$contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $user.Mail
							$itemsupdated += "EmailAddress"
						}
						$contact.Update("AutoResolve")
						
						Write-Host -ForegroundColor Green "Contact " -NoNewline
						Write-Host -ForegroundColor White $($contact.DisplayName) -NoNewline
						Write-Host -ForegroundColor Green " updated " -NoNewline
						Write-Host -ForegroundColor White "$($itemsupdated -join ', ')" -NoNewline
						Write-Host -ForegroundColor Green " successfully."
						
						$updatecount++
					} catch {
						$updatefailed += "$($contact.DisplayName)(Write)"
					}
				}
			}
		}
		
		Write-Host ""
		Write-Host -ForegroundColor White ($text = "Contacts update for mailbox $mailbox ($mailboxcount of $($TargetMailboxes.Count))")
		for ($i = 0; $i -lt $text.Length; $i++) {
			Write-Host -ForegroundColor White '-' -NoNewline
		}
		Write-Host ""; Write-Host ""
		Write-Host -ForegroundColor Green "Successful updates - $updatecount"
		if ($updatefailed) {
			Write-Host -ForegroundColor Red "Failed updates - $($updatefailed -join ', ')"
		}
		Write-Host -ForegroundColor White "Total Contacts attempted - $($updatecount + $updatefailed.Count)"
		Write-Host ""
		
		$emailtext += @"

	$mailbox ($mailboxcount of $($TargetMailboxes.Count))
	Contacts Updated Successfully
	$updatecount
	
"@
	
	if ($updatefailed) {
		$emailtext += @"
Contacts Failed to Update
$($updatefailed -join "`r`n")

"@
	}
	
	$mailboxcount++
}

Send-MailMessage -From 'Contacts Update <it@gopaschal.com>' -To 'Michael M Carter <mcarter@gopaschal.com>' -Subject "Contacts Update ($(Get-Date -Format 'MM/dd/yyyy'))" -Body $emailtext -DeliveryNotificationOption OnFailure, OnSuccess -SmtpServer 'mail.paschalcorp.com'

}

exit
