@{
	RootModule	      = '.\Paschal_Functions.psm1'
	ModuleVersion	  = '1.0.0.0'
	Author		      = 'Paschal IT'
	CompanyName	      = 'Paschal'
	FunctionsToExport = @("PSleep", "PPS", "ADSync", "PListSelect", "PSelect", "PDebug", "PInput", "PTitle", "Get-InstalledSoftware", "Enable-PaschalEXCContacts", "Update-PaschalEXCContacts", "Get-PaschalEXCContacts", "Update-PaschalFunctions", "Rename-PaschalComputer", "Remove-PaschalEXCContacts", "IsAdmin", "Update-PaschalWrightsoft")
	CmdletsToExport   = @()
	AliasesToExport   = @()
}