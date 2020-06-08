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
	PSleep 60 Syncing "Synchronizing with Office365"
	Write-Host ""
	Write-Host -ForegroundColor Green "Done!"
	Write-Host ""
}

function PListSelect {
 # Must pass in array of Strings!  Function returns an array with the selected strings.  Format - PListSelect [String Array] [Limit]
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
		[string[]]$list,
		
		[int]$limit = 0,
		
		[string]$prompt = $null,
		
		[boolean]$pclear = $false)
	
	$select = @(); $ret = @(); $offset = 0
	
	if ($list.Length -eq 0) {
		if (!$pclear) {
			Clear-Host
		}
		Write-Host -ForegroundColor Red "No input array provided.  Array or multiple values required.  Press enter to continue."
		Read-Host
		break
	}
	
	foreach ($i in $list) {
		$select += $false
	}
	
	do {
		if (!$pclear) {
			Clear-Host
		}
		# $count = 1 # Trash variable to track numbering
		
		if ($prompt) {
			Write-Host -ForegroundColor Cyan $prompt
			Write-Host ""
		}
		
		for ($i = 0; $i -lt 25; $i++) {
			# Write out options in columns of 15
			if (!$list[($i + $offset)]) {
				break
			}
			
			Write-Host -NoNewline ($i + $offset + 1); Write-Host -NoNewline ")`t"
			if ($select[($i + $offset)]) {
				Write-Host -ForegroundColor Green (($list[$i + $offset].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
			} else {
				Write-Host (($list[$i + $offset].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
			}
			if ($list[($i + $offset)].Length -gt 20) {
				Write-Host "..." -NoNewLine # If more than 20 characters, adds ellipses
			}
			
			if ($list[($i + $offset + 25)]) {
				# Checks to see if next column is needed; prints items 26-50 if they exist
				Write-Host -NoNewLine "`t"
				if ($list[($i + $offset)].Length -lt 8) {
					# These are to properly align everything in columns
					Write-Host -NoNewline "`t"
				}
				if ($list[($i + $offset)].Length -lt 16) {
					Write-Host -NoNewLine "`t"
				}
				Write-Host -NoNewline ($i + $offset + 26); Write-Host -NoNewline ")`t"
				if ($select[($i + $offset + 25)]) {
					Write-Host -ForegroundColor Green (($list[($i + $offset + 25)].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
				} else {
					Write-Host (($list[($i + $offset + 25)].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
				}
				if ($list[($i + $offset + 25)].Length -gt 20) {
					Write-Host "..." -NoNewLine # If more than 20 characters, adds ellipses
				}
			}
			
			if ($list[($i + $offset + 50)]) {
				# Checks to see if next column is needed; prints items 51-75 if they exist
				Write-Host -NoNewLine "`t"
				if ($list[($i + $offset + 25)].Length -lt 8) {
					# These are to properly align everything in columns
					Write-Host -NoNewLine "`t"
				}
				if ($list[($i + $offset + 25)].Length -lt 16) {
					Write-Host -NoNewLine "`t"
				}
				Write-Host -NoNewLine ($i + $offset + 51); Write-Host -NoNewLine ")`t"
				if ($select[($i + $offset + 50)]) {
					Write-Host -ForegroundColor Green (($list[($i + $offset + 50)].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
				} else {
					Write-Host (($list[($i + $offset + 50)].ToCharArray() | Select-Object -first 20) -join '') -NoNewline # If more than 20 characters, truncates to 20
				}
				if ($list[($i + $offset + 50)].Length -gt 20) {
					Write-Host "..." -NoNewLine # If more than 20 characters, adds ellipses
				}
			}
			
			Write-Host ""
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
		
		Write-Host -ForegroundColor Cyan "`r`nPlease select an item.  Selecting a highlighted item will deselect it.  Use 'A' to select all, or 'D' to deselect all.  Enter 'Y' when finished." -NoNewLine
		if ($limit) {
			Write-Host -ForegroundColor Red "  Limit of $limit selections." -NoNewLine
		}
		if ($list.Length -gt 75) {
			Write-Host ""
			if (($offset + 75) -lt $list.Length) {
				Write-Host -ForegroundColor Cyan "Use 'N' to display the next page.  " -NoNewLine
			}
			if (($offset - 75) -ge 0) {
				Write-Host -ForegroundColor Cyan "Use 'P' to display the previous page." -NoNewLine
			}
		}
		$x = Read-Host
		
		try {
			$x = [int]$x # Attempts to parse input to int.  Does nothing if input is not numerical.
		} catch [System.Management.Automation.PSInvalidCastException] {
		} # Prevent parse error from displaying.  It does not effect anything in the code.
		
		if ($x -match "^\d+$" -and $x -gt 0 -and $x -le $list.Length) {
			$x -= 1 # Make variable match array values
			
			if ($select[$x]) {
				# If true set false, and vice versa
				$select[$x] = $false
			} else {
				if (($select | Where-Object -FilterScript {
							$_ -eq $true
						}).Count -lt $limit -or $limit -eq 0) {
					$select[$x] = $true
				} else {
					Write-Host -ForegroundColor Red "`r`nCan't make selection as it exceeds the set limit of $limit items.  Please press enter and deselect one before choosing another."
					Read-Host | Out-Null
				}
			}
		} elseif ($x -eq 'd') {
			# Deselect all
			for ($i = 0; $i -lt $select.Length; $i++) {
				$select[$i] = $false
			}
		} elseif ($x -eq 'a') {
			# Select all
			if ($limit -ge $list.Length -or $limit -eq 0) {
				for ($i = 0; $i -lt $select.Length; $i++) {
					$select[$i] = $true
				}
			} else {
				Write-Host -ForegroundColor Red "`r`nCan't select all values as it exceeds the set limit of $limit items.  Please press enter and select individual values."
				Read-Host | Out-Null
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
			Write-Host -ForegroundColor Red "`r`nInput outside of available selection range.  Please press enter and try again."
			Read-Host | Out-Null
		}
	} while ($x -ne 'y') # Repeat until user keys Y
	
	for ($i = 0; $i -lt $select.Length; $i++) {
		# Fill return array with selected strings
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
		[string[]]$list,
		
		[string]$prompt = $null,
		
		[boolean]$pclear = $false)
	
	if ($list.Length -le 0) {
		# Check that array exists and isn't empty
		if (!$pclear) {
			Clear-Host
		}
		Write-Host -ForegroundColor Red "Input list is empty or missing.  Please check your function call."
		Read-Host | Out-Null
		break
	}
	
	$check = $false
	
	do {
		if (!$pclear) {
			Clear-Host
		}
		$count = 1
		
		if ($prompt) {
			# Print custom prompt if it exists, else print nothing
			Write-Host ""
			Write-Host -ForegroundColor Cyan $prompt
			Write-Host ""
		}
		
		foreach ($i in $list) {
			# Print list of options with numbering
			Write-Host -NoNewLine "$count)`t"
			Write-Host $i
			
			$count += 1
		}
		
		Write-Host "`r`n99)`tBack/Exit/Cancel"
		
		Write-Host -ForegroundColor Cyan "`r`nPlease select an item.  " -NoNewLine
		$x = Read-Host
		
		try {
			$x = [int]$x
		} catch {
		}
		
		if ($x -match "^\d+$" -and (($x -gt 0 -and $x -le $list.Length) -or $x -eq 99)) {
			# Make sure user input is number and within range
			return ($x - 1) # Returns user input minus 1 so it matches proper array value
		} else {
			# Error if invalid input
			Write-Host -ForegroundColor Red "`r`nInvalid selection.  Please press enter and try again."
			Read-Host | Out-Null
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
		[string]$prompt,
		
		[int]$req = -1,
		
		[int]$allownull = 0)
	
	if ($req -eq -1 -and !$allownull) {
		$allownull = 1
	}
	
	do {
		Write-Host $prompt
		$var = Read-Host
		
		if (($allownull -and $var.Length -eq 0) -or ($var.Length -eq $req -and $req -gt 0) -or $req -eq -1 -or ($req -eq 0 -and $var.Length -gt 0)) {
			# if (null allowed and input null) (requirement exists and input matches requirement) (no requirement, so anything goes) (requirement is 0 so input not allowed, input exists and is not null)
			return $var # If required conditions are met, return the input
		} elseif (!$var) {
			Write-Host ""
			if (!$Script:pclear) {
				Clear-Host
			}
			Write-Host -ForegroundColor Red "Input required.  Please try again.`r`n"
		} else {
			Write-Host ""
			if (!$Script:pclear) {
				Clear-Host
			}
			Write-Host -ForegroundColor Red "Required length not met.  Please try again.`r`n"
		}
	} while (2 -lt 3) # Repeat indefinitely until return conditions are met
}

function PTitle {
	param ([string]$title = "Paschal IT",
		
		[string]$version = "0.0")
	
	$Host.UI.RawUI.WindowTitle = "$title v$version"
}

filter Get-InstalledSoftware {
 ## Courtesy of Chris Dent, Powershell Guru
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
							Name            = $name
							DisplayName     = $packageKey.GetValue('DisplayName')
							DisplayVersion  = $packageKey.GetValue('DisplayVersion')
							InstallDate     = $installDate
							InstallLocation = $packageKey.GetValue('InstallLocation')
							HelpLink        = $packageKey.GetValue('HelpLink')
							Publisher       = $packageKey.GetValue('Publisher')
							UninstallString = $packageKey.GetValue('UninstallString')
							URLInfoAbout    = $packageKey.GetValue('URLInfoAbout')
							Is64Bit         = $is64Bit
							Hive            = $baseKey.Name
							Path            = Join-Path $key $name
							Username        = $username
							ComputerName    = $ComputerName
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
	<#
	.SYNOPSIS
		Adds all existing Paschal contacts to a specified Exchange mailbox.
	
	.DESCRIPTION
		Adds all existing Paschal contacts from Active Directory to a specified Exchange mailbox.
	
	.PARAMETER TargetMailbox
		(Required) Email Address of the Exchange mailbox to receive new contacts.
	
	.PARAMETER Credentials
		(Optional) Exchange credentials with permissions to Read, Write, and Impersonate.  If not supplied via parameters, the user will be prompted to input credentials in order to complete the command.
	
	.EXAMPLE
		PS C:\> Enable-PaschalEXCContacts -TargetMailbox email@gopaschal.com -Credentials $myCred
#>
	
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
	
	$userlist = Get-ADUser -Filter * -SearchBase "OU=Springdale, DC=US, DC=PaschalCorp, DC=com" -Properties DisplayName, Mobile, TelephoneNumber, Mail, GivenName, Surname, Department, Title, ProxyAddresses | Where-Object {
		$_.Mobile -or $_.TelephoneNumber -or $_.SamAccountName -eq "_temp" -or $_.SamAccountName -eq "warehousenight"
	}
	
	$contacthash = $userlist |
		Where-Object {
			$_.DistinguishedName -match "OU=Users" -or $_.Name -eq "CSR On Call" -or $_.SamAccountName -eq "warehousenight"
		} |
		Select-PSFObject @(
			"GivenName as FirstName"
			"Surname as LastName"
			"DisplayName"
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
					# 
					New-EXCContact -MailboxName $targetUser.Mail -MailboxOwnerDistinguishedName $targetUser.DistinguishedName -FirstName $contact.FirstName -LastName $contact.LastName -DisplayName $contact.DisplayName -BusinessPhone $contact.BusinessPhone -MobilePhone $contact.MobilePhone -EmailAddress $contact.EmailAddress -Department $contact.Department -JobTitle $contact.Title -CompanyName "Paschal Air, Plumbing & Electric - AR01-PS" -Credentials $exchcred -useImpersonation
					Write-Host -ForegroundColor Green "Complete!"
					# Start-Sleep -Seconds 1
					Write-Host -ForegroundColor Yellow "Retrieving Contact..."
					$temp = Get-PaschalEXCContacts -MailboxName $targetUser.Mail -Credentials $exchcred | Where-Object {
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
	<#
	.SYNOPSIS
		Updates Paschal controlled properties of Paschal contacts on all company phones.
	
	.DESCRIPTION
		Updates Paschal controlled properties (Display Name, Primary Email, Mobile Phone, Business Phone, SID) of all Paschal contacts with "AR01-PS" in the Company Name on all company phones, as listed in Active Directory.  If missing contacts are found, the script will attempt to create a contact to replace them.
	
	.PARAMETER TargetMailboxes
		(Optional) A string or array of strings containing Email Addresses of specific Exchange mailboxes to be targeted for updates.  If not supplied via parameter, the script will run for all active mailboxes on the Paschal domain.
	
	.PARAMETER Credentials
		(Optional) Exchange server credentials with permissions to Read, Write, and Impersonate.  If not supplied via parameter, the user will be prompted to input credentials in order to complete the command.
	
	.PARAMETER TargetContacts
		A description of the TargetContacts parameter.
	
	.EXAMPLE
		PS C:\> Update-PaschalEXCContacts -TargetMailboxes "email@gopaschal.com" -Credentials $myCred
	
	.EXAMPLE
		PS C:\> Update-PaschalEXCContacts -Credentials $myCred
	
	.NOTES
		Additional information about the function.
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $false,
			ValueFromPipeline = $true)]
		[ValidatePattern('@gopaschal\.com|@paschalcorp\.com')]
		[string[]]$TargetMailboxes = @((Get-ADUser -Filter * -SearchBase "OU=Users, OU=Springdale, DC=US, DC=PaschalCorp, DC=com" -Properties Mail, Mobile | Where-Object {
					$_.Mobile
				}).Mail + (Get-ADUser -Identity csr-on-call -Properties Mail).Mail + (get-aduser -Identity warehousenight -Properties Mail).Mail),
		
		[Parameter(Mandatory = $false)]
		[PSCredential]$Credentials = (Get-Credential srv),
		
		[Parameter(Mandatory = $false)]
		[string[]]$TargetContacts,

		[Parameter(Mandatory = $false)]
		[switch]$NoEmail
	)
	
	$ErrorActionPreference = 'Continue'
	
	# $dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
	# [void][Reflection.Assembly]::LoadFile($dllpath)
	
	$exchcred = $Credentials
	if (-not $exchcred) {
		$exchcred = Get-Credential srv
	}
	
	if ($TargetContacts) {
		$userlist = $TargetContacts | ForEach-Object {
			Get-ADUser -Filter "ANR -eq '$_'" -SearchBase "OU=Springdale, DC=US, DC=PaschalCorp, DC=com" -Properties DisplayName, Mobile, TelephoneNumber, Mail, GivenName, Surname, Department, Title, ProxyAddresses | Where-Object {
				$_.DistinguishedName -notmatch "Terminated"
			}
		}
	} else {
		$userlist = Get-ADUser -Filter * -SearchBase "OU=Springdale, DC=US, DC=PaschalCorp, DC=com" -Properties DisplayName, Mobile, TelephoneNumber, Mail, GivenName, Surname, Department, Title, ProxyAddresses | Where-Object {
			$_.Mobile -or $_.TelephoneNumber -or $_.SamAccountName -eq "_temp" -or $_.SamAccountName -eq "warehousenight"
		}
	}
	
	$contacthash = $userlist |
		Where-Object {
			$_.DistinguishedName -match "OU=Users" -or $_.SamAccountName -eq "csr-on-call" -or $_.SamAccountName -eq "warehousenight"
		} |
		Select-PSFObject @(
			"GivenName as FirstName"
			"Surname as LastName"
			"DisplayName"
			"Mobile as MobilePhone"
			"TelephoneNumber as BusinessPhone"
			"Mail as EmailAddress"
			"SID as Mileage"
			"Title"
			"Department"
		) |
		ConvertTo-PSFHashtable -Include FirstName, LastName, BusinessPhone, MobilePhone, DisplayName, Department, Title, EmailAddress, Mileage
	$emailtext = $null; $mailboxcount = 1;
	
	Clear-Host
	foreach ($mailbox in $TargetMailboxes) {
		$updatecount = 0; $updatefailed = @(); $contactlist = @(); $updateadded = @();
		
		Write-Host -ForegroundColor White ($text = "Current Target - $mailbox ($mailboxcount of $($TargetMailboxes.Count))")
		for ($i = 0; $i -lt $text.Length; $i++) {
			Write-Host -ForegroundColor White '-' -NoNewline
		}
		Write-Host ""
		
		$contactlist = Get-PaschalEXCContacts -MailboxName $mailbox -Credentials $exchcred
		
		if (-not $contactlist) {
			$emailtext += @"
$mailbox
Could not update Contacts
Unable to locate Contacts on Exchange Mailbox

"@
		} else {
			foreach ($contact in $contacthash) {
				$contactmatch = $null;
				
				$contactmatch = $contactlist | Where-Object {
					$_.Mileage -eq $contact.Mileage
				}
				
				if (-not $contactmatch) {
					$contactmatch = $contactlist | Where-Object {
						$(
							$count = 0
							if ($_.GivenName -eq $contact.FirstName) {
								$count++
							}
							if ($_.Surname -eq $contact.LastName) {
								$count++
							}
							if ($_.DisplayName -eq $contact.DisplayName) {
								$count++
							}
							if ($_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] -eq $contact.MobilePhone) {
								$count++
							}
							if ($_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address -eq $contact.EmailAddress) {
								$count++
							}
							$count
						) -ge 3
					}
				}
				
				if (-not $contactmatch) {
					try {
						# -MailboxOwnerDistinguishedName (Get-ADUser -Filter "ANR -eq '$mailbox'").DistinguishedName
						New-EXCContact -MailboxName $mailbox -FirstName $contact.FirstName -LastName $contact.LastName -DisplayName $contact.DisplayName -BusinssPhone $contact.BusinessPhone -MobilePhone $contact.MobilePhone -EmailAddress $contact.EmailAddress -Department $contact.Department -JobTitle $contact.Title -CompanyName "Paschal Air, Plumbing & Electric - AR01-PS" -Credentials $exchcred -useImpersonation
						Write-Host -ForegroundColor Green "Missing contact found - " -NoNewline
						Write-Host -ForegroundColor White $($contact.DisplayName) -NoNewline
						Write-Host -ForegroundColor Green " - Created Successfully"
						$updateadded += $contact.DisplayName
						try {
							$temp = Get-PaschalEXCContacts -MailboxName $mailbox -EmailAddress $contact.EmailAddress -Credentials $exchcred
							$temp.Mileage = $contact.Mileage
							$temp.Update("AutoResolve")
							Write-Host -ForegroundColor Green "Set SID as Mileage successfully"
						} catch {
							Write-Host -ForegroundColor Yellow "Could not set SID as Mileage"
						}
						$updatecount++
					} catch {
						Write-Host -ForegroundColor Red "Contact creation failed for missing contact - " -NoNewline
						Write-Host -ForegroundColor White $contact.DisplayName
						$updatefailed += "$($contact.DisplayName)(Creation)"
					}
					
				} else {
					try {
						$itemsupdated = @(); $itemsvalidated = @()
						if ($contactmatch.Mileage -eq $contact.Mileage) {
							$itemsvalidated += "SID"
						} else {
							$contactmatch.Mileage = $contact.Mileage
							$itemsupdated += "SID"
						}
						if ($contactmatch.DisplayName -eq $contact.DisplayName) {
							$itemsvalidated += "DisplayName"
						} else {
							$contactmatch.DisplayName = $contact.DisplayName
							$itemsupdated += "DisplayName"
						}
						if ($contact.MobilePhone -and $contactmatch.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] -eq $contact.MobilePhone) {
							$itemsvalidated += "MobilePhone"
						} elseif ($contact.MobilePhone) {
							$contactmatch.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $contact.MobilePhone
							$itemsupdated += "MobilePhone"
						}
						if ($contact.BusinessPhone -and $contactmatch.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] -eq $contact.BusinessPhone) {
							$itemsvalidated += "BusinessPhone"
						} elseif ($contact.BusinessPhone) {
							$contactmatch.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $contact.BusinessPhone
							$itemsupdated += "BusinessPhone"
						}
						if ($contact.EmailAddress -and $contactmatch.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address -eq $contact.EmailAddress) {
							$itemsvalidated += "EmailAddress"
						} elseif ($contact.EmailAddress) {
							$contactmatch.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $contact.EmailAddress
							$itemsupdated += "EmailAddress"
						}
						$contactmatch.Update("AutoResolve")
						
						Write-Host -ForegroundColor Green "Contact " -NoNewline
						Write-Host -ForegroundColor White $contactmatch.DisplayName -NoNewline
						if ($itemsvalidated) {
							Write-Host -ForegroundColor Green " validated " -NoNewline
							Write-Host -ForegroundColor White "$($itemsvalidated -join ', ')" -NoNewline
						}
						if ($itemsvalidated -and $itemsupdated) {
							Write-Host -ForegroundColor White "," -NoNewline
						}
						if ($itemsupdated) {
							Write-Host -ForegroundColor Yellow " updated " -NoNewline
							Write-Host -ForegroundColor White "$($itemsupdated -join ', ')" -NoNewline
						}
						Write-Host -ForegroundColor Green " successfully."
						
						$updatecount++
					} catch {
						$updatefailed += "$($contactmatch.DisplayName)(Write)"
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
		Write-Host -ForegroundColor Green "Contacts successfully validated - $updatecount"
		if ($updatefailed) {
			Write-Host -ForegroundColor Red "Contacts failed to validate - $($updatefailed -join ', ')"
		}
		Write-Host -ForegroundColor White "Total Contacts attempted - $($updatecount + $updatefailed.Count)"
		Write-Host ""
		
		$emailtext += @"

	$mailbox ($mailboxcount of $($TargetMailboxes.Count))
	Contacts Validated Successfully
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
	
	if (-not $NoEmail) {
		Send-MailMessage -From 'Contacts Update <it@gopaschal.com>' -To 'Paschal IT <it@gopaschal.com>' -Subject "Contacts Update ($(Get-Date -Format 'MM/dd/yyyy'))" -Body $emailtext -DeliveryNotificationOption OnFailure, OnSuccess -SmtpServer 'mail.paschalcorp.com'
	}
}


function Get-PaschalEXCContacts {
	<#
	.SYNOPSIS
		Retrieves contacts from an Exchange mailbox within the Paschal domain.
	
	.DESCRIPTION
		Retrieves contacts from a specified Exchange mailbox within the Paschal domain, with options to filter by Name (First, Last, Display) or Email Address.  If no filter is specified, Get-PaschalEXCContacts returns all contacts from the mailbox.
	
	.PARAMETER MailboxName
		(Required) Email address of the Exchange mailbox from which to retrieve contacts.
	
	.PARAMETER Credentials
		(Optional) Exchange server credentials with permissions to Read, Write, and Impersonate.  If not supplied, the user will be prompted for the required credentials in order to complete the command.
	
	.PARAMETER EmailAddress
		(Optional) Email Address of target contact with which to filter results.
	
	.PARAMETER FirstName
		(Optional) First Name of target contact with which to filter results.  May be used singularly or paired with LastName.
	
	.PARAMETER LastName
		(Optional) Last Name of target contact with which to filter results.  May be used singularly or paired with FirstName.
	
	.PARAMETER DisplayName
		(Optional) Full Display Name of target contact with which to filter results.
	
	.PARAMETER All
		(Optional) If $true, returns all contacts from the specified Exchange mailbox.  If $false, refines results to only those with Company Name matching "AR01-PS".
	
	.EXAMPLE
		PS C:\> Get-PaschalEXCContacts -MailboxName email@domain.com -Credentials $myCred -EmailAddress contact@domain.com
	
		Retrieves contacts with email addresses matching "contact@domain.com" from the Exchange mailbox "email@domain.com".
	
	.EXAMPLE
		PS C:\> Get-PaschalEXCContacts -MailboxName email@domain.com
	
		Retrieves ALL contacts with Company Name matching "AR01-PS" from the Exchange mailbox "email@domain.com", prompting for credentials.
	
	.EXAMPLE
		PS C:\> Get-PaschalEXCContacts -MailboxName email@domain.com -Credentials $myCred -All
	
		Retrieves ALL contacts from the Exchange mailbox "email@domain.com".
#>
	
	[CmdletBinding(DefaultParameterSetName = 'Default')]
	param
	(
		[Parameter(ParameterSetName = 'EmailAddress',
			Mandatory = $true,
			ValueFromPipeline = $true,
			Position = 1)]
		[Parameter(ParameterSetName = 'FirstLastName',
			Mandatory = $true,
			ValueFromPipeline = $true,
			Position = 1)]
		[Parameter(ParameterSetName = 'DisplayName',
			Mandatory = $true,
			ValueFromPipeline = $true,
			Position = 1)]
		[Parameter(ParameterSetName = 'Default',
			Position = 1)]
		[ValidatePattern('@gopaschal\.com|@paschalcorp\.com')]
		[string]$MailboxName,
		
		[Parameter(ParameterSetName = 'EmailAddress')]
		[Parameter(ParameterSetName = 'FirstLastName')]
		[Parameter(ParameterSetName = 'DisplayName')]
		[Parameter(ParameterSetName = 'Default')]
		[pscredential]$Credentials = (Get-Credential srv),
		
		[Parameter(ParameterSetName = 'EmailAddress',
			Position = 2)]
		[string]$EmailAddress,
		
		[Parameter(ParameterSetName = 'FirstLastName',
			Position = 2)]
		[string]$FirstName,
		
		[Parameter(ParameterSetName = 'FirstLastName',
			Position = 3)]
		[string]$LastName,
		
		[Parameter(ParameterSetName = 'DisplayName',
			Position = 2)]
		[string]$DisplayName,
		
		[Parameter(ParameterSetName = 'EmailAddress')]
		[Parameter(ParameterSetName = 'FirstLastName')]
		[Parameter(ParameterSetName = 'DisplayName')]
		[Parameter(ParameterSetName = 'Default')]
		[switch]$All = $false
	)
	
	begin {
		try {
			$contacts = Get-EXCContacts -MailboxName $MailboxName -Credentials $Credentials -useImpersonation
			if (-not $All) {
				$contacts = $contacts | Where-Object {
					$_.CompanyName -match "AR01-PS"
				}
			}
		} catch {
			throw "Contacts could not be retrieved for Mailbox $MailboxName.`r`nPlease ensure you are inputting a valid Paschal email address and try again."
		}
	}
	process {
		switch ($PsCmdlet.ParameterSetName) {
			'EmailAddress' {
				
				$contacts = $contacts | Where-Object {
					$_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address -eq $EmailAddress
				}
				
				break
			}
			'FirstLastName' {
				
				if ($FirstName) {
					$contacts = $contacts | Where-Object {
						$_.GivenName -eq $FirstName
					}
				}
				if ($LastName) {
					$contacts = $contacts | Where-Object {
						$_.Surname -eq $LastName
					}
				}
				
				break
			}
			'DisplayName' {
				
				$contacts = $contacts | Where-Object {
					$_.DisplayName -eq $DisplayName
				}
				
				break
			}
		}
		
	}
	end {
		if ($contacts) {
			return $contacts
		} else {
			Write-Host ""
			Write-Host -ForegroundColor Red "No contact was found in Mailbox $MailboxName matching search criteria."
			Write-Host -ForegroundColor Red "Please double check your search criteria and try again."
			Write-Host ""
			return $null
		}
	}
}


function Update-PaschalFunctions {

	<#
	.SYNOPSIS
		Pulls newest version of Paschal_Functions module and imports it.
	.DESCRIPTION
		Copies Paschal_Functions.psm1 and Paschal_Functions.psd1 from fileserver, removes current Paschal_Functions module, and imports new Paschal_Functions module.
	.EXAMPLE
		PS C:\> Update-PaschalFunctions
		This function has no parameters.  When used, it will copy the newest version of Paschal_Functions and replace the existing module with the new version.
	#>

	[CmdletBinding()]
	param ()
	
	Copy-Item '\\wfs01v\Paschal$\Deployment\Reference\Replace\Paschal_Functions\Paschal_Functions.psm1' C:\Paschal\Reference\Paschal_Functions\Paschal_Functions.psm1 -Force
	Copy-Item '\\wfs01v\Paschal$\Deployment\Reference\Replace\Paschal_Functions\Paschal_Functions.psd1' C:\Paschal\Reference\Paschal_Functions\Paschal_Functions.psd1 -Force
	
	Remove-Module Paschal_Functions
	Import-Module Paschal_Functions -Global
	
}

function Rename-PaschalComputer {
	
	<#
	.SYNOPSIS
		Renames target computer with a new name.
	.DESCRIPTION
		Renames target computer (ComputerName) with a new name (NewName) in Windows and Active Directory (if computer is on a domain).
	.EXAMPLE
		PS C:\> Rename-PaschalComputer -ComputerName $currentName -NewName $newName -Credentials $myCred
		Renames the computer detailed in $currentName to the name detailed in $newName.

	.EXAMPLE
		PS C:\> Rename-PaschalComputer -ComputerName $currentName -NewName $newName -Restart
		Prompts the user for credentials since none were passed via parameter, then renames the computer detailed in $currentName to the name detailed in $newName, then restarts the target computer.

	.PARAMETER ComputerName
		(Required) Full name of computer to be renamed.

	.PARAMETER NewName
		(Optional) Full name for computer to be renamed to.

	.PARAMETER Restart
		(Optional) Defaults to false.  If true, will restart the target computer after rename.

	.PARAMETER Credentials
		(Optional) Credentials with privileges to make changes on the current domain.

	#>

	[CmdletBinding()]
	param(
		[Parameter()]
		[string]$ComputerName,

		[Parameter()]
		[string]$NewName,

		[Parameter()]
		[switch]$Restart,

		[Parameter()]
		[PSCredential]$Credentials
	)

	if (-not $ComputerName) {
		$ComputerName = Read-Host -Prompt "Please input the name of the Target Computer, if different from $env:COMPUTERNAME"
		if (-not $ComputerName) {
			$ComputerName = $env:COMPUTERNAME
		}
	}

	if (-not $NewName) {
		$NewName = Read-Host -Prompt "Please input the New Name for $ComputerName"
	}

	if ($Credentials -isnot [pscredential]) {
		if (-not $Credentials) {
			$Credentials = Get-Credential
		} else {
			$Credentials = Get-Credential $Credentials
		}
	}

	Invoke-Command -ComputerName WDC01V -Credential $Credentials -ScriptBlock {
		if (-not (Get-ADComputer -Identity $args[0] -Properties MemberOf).MemberOf -match "CN=Weekly-Reboot,OU=ComputersGroups,OU=Groups,OU=Springdale,DC=US,DC=PaschalCorp,DC=com") {
			Add-ADPrincipalGroupMembership -Identity $args[0] -MemberOf "CN=Weekly-Reboot,OU=ComputersGroups,OU=Groups,OU=Springdale,DC=US,DC=PaschalCorp,DC=com" -Confirm:$false
		}
	} -ArgumentList $ComputerName

	if ((Read-Host -Prompt "Remove computer from Weekly-Reboot AD Group?") -eq 'y') {
		Invoke-Command -ComputerName WDC01V -Credential $Credentials -ScriptBlock {
			Remove-ADPrincipalGroupMembership -Identity $args[0] -MemberOf "CN=Weekly-Reboot,OU=ComputersGroups,OU=Groups,OU=Springdale,DC=US,DC=PaschalCorp,DC=com" -Confirm:$false -ErrorAction SilentlyContinue
		} -ArgumentList $ComputerName
	}

	if (-not $Restart) {
		if ((Read-Host -Prompt "Restart Target Computer when rename is complete?") -eq 'y') {
			$Restart = $true
		}
	}

	$splat = @{
		ComputerName     = $ComputerName
		NewName          = $NewName
		Restart          = $Restart
		DomainCredential = $Credentials
	}

	Write-Host ""
	Write-Host -ForegroundColor White "Renaming $ComputerName to $NewName.  Is this correct?"
	$confirm = Read-Host -Prompt "(Y/N)"

	if ($confirm -ne "y") {
		Write-Host ""
		Write-Host "Rename canceled.  Computer Name will remain $ComputerName."
		return
	} else {
		try {
			Rename-Computer @splat
		} catch {
			Write-Host ""
			Write-Host -ForegroundColor Red "Rename failed.  Computer Name will remain $ComputerName."
		}
	}

}

function Remove-PaschalEXCContacts {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		$TargetContact,

		[Parameter()]
		[ValidatePattern('@(gopaschal|paschalcorp)\.com')]
		[string[]]$TargetMailboxes,

		[Parameter()]
		[PSCredential]$Credentials
	)

	if ($Credentials -isnot [PSCredential] -or -not $Credentials) {
		$Credentials = Get-Credential
	}

	$TargetContact = @(Get-ADUser -Filter "ANR -eq '$TargetContact'" -Properties Mail, Mobile, TelephoneNumber) | Where-Object {
		$_.SamAccountName -notmatch "-adm"
	}

	if (-not $TargetContact) {
		throw "$TargetContact does not match any user information"
	} elseif ($TargetContact.Count -gt 1) {
		throw "$TargetContact returned multiple matches.  Please refine search and try again"
	} else {
		if (-not $TargetMailboxes) {
			$TargetMailboxes = @((Get-ADUser -Filter * -SearchBase "OU=Users, OU=Springdale, DC=US, DC=PaschalCorp, DC=com" -Properties Mail, Mobile | Where-Object {
						$_.Mobile
					}).Mail + (Get-ADUser -Identity csr-on-call -Properties Mail).Mail)
		}

		foreach ($mailbox in $TargetMailboxes) {
			Write-Host $mailbox
			$contacts = Get-PaschalEXCContacts -MailboxName $mailbox -Credentials $Credentials

			$contacts.Count

			$contact = $contacts | Where-Object {
				$_.Mileage -eq $TargetContact.SID
			}

			$contact.Count
			$contact.DisplayName

			if ($contact) {
				$contact.Delete("MoveToDeletedItems")
			} else {
				Write-Host "Could not locate single contact matching target."
			}
		}
	}

}
