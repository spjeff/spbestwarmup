﻿<#
.SYNOPSIS
	Warm up SharePoint IIS W3WP memory cache by loading pages from WebRequest

.DESCRIPTION
	Loads the full page so resources like CSS, JS, and images are included.  Please modify lines 331-345
	to suit your portal content design (popular URLs, custom pages, etc.)
	
	Comments and suggestions always welcome!  Please, use the issues panel at the project page.

.PARAMETER url
	A collection of url that will be added to the list of websites the script will fetch.
	
.PARAMETER install
	Typing "SPBestWarmUp.ps1 -install" will create a local Task Scheduler job under credentials of the
	current user.  Job runs every 60 minutes on the hour to help automatically populate cache.  
	Keeps cache full even after IIS daily recycle, WSP deployment, reboot, or other system events.

.PARAMETER installfarm
	Typing "SPBestWarmUp.ps1 -farminstall" will create a Task Scheduler job on all machines in the farm.

.PARAMETER uninstall
	Typing "SPBestWarmUp.ps1 -uninstall" will remove Task Scheduler job from all machines in the farm.
	
.PARAMETER skiplog
	Typing "SPBestWarmUp.ps1 -skiplog" will avoid writing to the EventLog.
	
.PARAMETER allsites
	Typing "SPBestWarmUp.ps1 -allsites" will load every site and web URL.
	
.EXAMPLE
	.\SPBestWarmUp.ps1 -url "http://domainA.tld","http://domainB.tld"

.EXAMPLE
	.\SPBestWarmUp.ps1 -i
	.\SPBestWarmUp.ps1 -install

.EXAMPLE
	.\SPBestWarmUp.ps1 -f
	.\SPBestWarmUp.ps1 -installfarm

.EXAMPLE
	.\SPBestWarmUp.ps1 -u
	.\SPBestWarmUp.ps1 -uninstall

	
.NOTES  
	File Name		:	SPBestWarmUp.ps1
	Author			:	Jeff Jones  - @spjeff
	Author			:	Hagen Deike - @hd_ka
	Version			:	2.2.4
	Modified		:	05-13-2016

.LINK
	https://github.com/spjeff/spbestwarmup
#>

[CmdletBinding()]
param (
	[Parameter(Mandatory=$False, Position=0, ValueFromPipeline=$false, HelpMessage='A collection of URLs that will be fetched too')]
	[Alias("url")]
	[ValidateNotNullOrEmpty()]
	[ValidatePattern("https?:\/\/\D+")]
	[string[]]$cmdurl,

	[Parameter(Mandatory=$False, Position=1, ValueFromPipeline=$false, HelpMessage='Use -install -i parameter to add script to Windows Task Scheduler on local machine')]
	[Alias("i")]
	[switch]$install,
	
	[Parameter(Mandatory=$False, Position=2, ValueFromPipeline=$false, HelpMessage='Use -installfarm -f parameter to add script to Windows Task Scheduler on all farm machines')]
	[Alias("f")]
	[switch]$installfarm,
	
	[Parameter(Mandatory=$False, Position=3, ValueFromPipeline=$false, HelpMessage='Use -uninstall -u parameter to remove Windows Task Scheduler job')]
	[Alias("u")]
	[switch]$uninstall,
	
	[Parameter(Mandatory=$False, Position=4, ValueFromPipeline=$false, HelpMessage='Use -skiplog -sl parameter to avoid writing to Event Log')]
	[Alias("sl")]
	[switch]$skiplog,
	
	[Parameter(Mandatory=$False, Position=5, ValueFromPipeline=$false, HelpMessage='Use -allsites -all parameter to load every site and web')]
	[Alias("all")]
	[switch]$allsites
)

Function Installer() {
	# Add to Task Scheduler
	Write-Output "  Installing to Task Scheduler..."
	$user = $ENV:USERDOMAIN + "\"+$ENV:USERNAME
	Write-Output "  Current User: $user"
	
	# Attempt to detect password from IIS Pool (if current user is local admin and farm account)
	$appPools = Get-CimInstance -Namespace "root/MicrosoftIISv2" -ClassName "IIsApplicationPoolSetting" -Property Name, WAMUserName, WAMUserPass | Select-Object WAMUserName, WAMUserPass
	foreach ($pool in $appPools) {			
		if ($pool.WAMUserName -like $user) {
			$pass = $pool.WAMUserPass
			if ($pass) {
				break
			}
		}
	}
	
	# Manual input if auto detect failed
	if (!$pass) {
		$pass = Read-Host "Enter password for $user "
	}
	
	# Task Scheduler command
	if ($allsites) {$suffix += " -allsites"}
	if ($skiplog) {$suffix += " -skiplog"}
	$cmd = """PowerShell.exe -ExecutionPolicy Bypass '$cmdpath$suffix'"""
	
	# Target machines
	$machines = @()
	if ($installfarm -or $uninstall) {
		# Create farm wide on remote machines
		foreach ($srv in (Get-SPServer | Where-Object {$_.Role -ne "Invalid"})) {
			$machines += $srv.Address
		}
	} else {
		# Create local on current machine
		$machines += "localhost"
	}
	$machines | ForEach-Object {
		if ($uninstall) {
			# Delete task
			Write-Output "SCHTASKS DELETE on $_"
			schtasks /s $_ /delete /tn "SPBestWarmUp" /f
			Write-Host "  [OK]" -Fore Green
		} else {
			# Copy local file to remote UNC path machine
			Write-Output "SCHTASKS CREATE on $_"
			if ($_ -ne "localhost" -and $_ -ne $ENV:COMPUTERNAME) {
				$dest = $cmdpath
				$drive = $dest.substring(0,1)
				$match =  Get-CimInstance -Class Win32_LogicalDisk | Where-Object {$_.DeviceID -eq ($drive+":") -and $_.DriveType -ne 4}
				if ($match) {
					$dest = "\\" + $_ + "\" + $drive + "$" + $dest.substring(2,$dest.length-2)
					mkdir (Split-Path $dest) -ErrorAction SilentlyContinue | Out-Null
					Write-Output $dest
					Copy-Item $cmdpath $dest -Confirm:$false
				}
			}
			# Create task
			schtasks /s $_ /create /tn "SPBestWarmUp" /ru $user /rp $pass /rl highest /sc daily /st 01:00 /ri 60 /du 24:00 /tr $cmd /f
			Write-Host "  [OK]" -Fore Green
		}
	}
}

Function WarmUp() {
	# Load plugin
	Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

	# Warm up CMD parameter URLs
	$cmdurl | ForEach-Object {NavigateTo $_}

	# Warm up SharePoint web applications
	Write-Output "Opening Web Applications..."

	# Accessing the Alternate URl Collection makes sure to warm up all "extended webs" (in this case existing multiple IIS webapps exists for one SharePoint webapp)
	# If the SharePoint Webs are not extended, then this creates additional unnecessary, but quick iterations, since the webapps are already warmed up under another Alternate URL.
	$altUrls = (Get-SPAlternateURL)
	foreach ($altUrl in $altUrls) {
		$url = $altUrl.PublicUrl
        	$wa = (Get-SPWebApplication $url -IncludeCentralAdministration | Sort-Object IsAdministrationWebApplication)
	        
		NavigateTo $url
		NavigateTo $url"_api/web"
        	NavigateTo $url"_api/_trust" #in adfs environments the first user login is slow if this URL is not warmed up 
		NavigateTo $url"_layouts/viewlsts.aspx"
		NavigateTo $url"_vti_bin/UserProfileService.asmx"
		NavigateTo $url"_vti_bin/sts/spsecuritytokenservice.svc"

		if ($wa.IsAdministrationWebApplication) {
			# Central Admin
			NavigateTo $url"Lists/HealthReports/AllItems.aspx"
			NavigateTo $url"_admin/FarmServers.aspx"
			NavigateTo $url"_admin/Server.aspx"
			NavigateTo $url"_admin/WebApplicationList.aspx"
			NavigateTo $url"_admin/ServiceApplications.aspx"
			
			# Manage Service Application
			$sa = Get-SPServiceApplication
			$links = $sa | ForEach-Object {$_.ManageLink.Url} | Select-Object -Unique
			foreach ($link in $links) {
				$ml = $link.TrimStart('/')
				NavigateTo "$url$ml"
			}
		}		
		if ($allsites) {
			# Warm Up Individual Site Collections and Sites
 			$sites = (Get-SPSite -WebApplication $wa -Limit ALL)
 			foreach($site in $sites){
 				$webs = (Get-SPWeb -Site $site -Limit ALL)
 				foreach($web in $webs){
 					$url = $web.Url
 					NavigateTo $url
 				}
 			}
		}
	}
	
	# Warm up Service Applications
	Get-SPServiceApplication | ForEach-Object {$_.EndPoints | ForEach-Object {$_.ListenUris | ForEach-Object {NavigateTo $_.AbsoluteUri}}}

	# Warm up Project Server
	Write-Output "Opening Project Server PWAs..."
	if ((Get-Command Get-SPProjectWebInstance -ErrorAction SilentlyContinue).Count -gt 0) {
		Get-SPProjectWebInstance | ForEach-Object {
			# Thanks to Eugene Pavlikov for the snippet
			$url = ($_.Url).AbsoluteUri + "/"
		
			NavigateTo $url
			NavigateTo ($url + "_layouts/viewlsts.aspx")
			NavigateTo ($url + "_vti_bin/UserProfileService.asmx")
			NavigateTo ($url + "_vti_bin/sts/spsecuritytokenservice.svc")
			NavigateTo ($url + "Projects.aspx")
			NavigateTo ($url + "Approvals.aspx")
			NavigateTo ($url + "Tasks.aspx")
			NavigateTo ($url + "Resources.aspx")
			NavigateTo ($url + "ProjectBICenter/Pages/Default.aspx")
			NavigateTo ($url + "_layouts/15/pwa/Admin/Admin.aspx")
		}
	}

	# Warm up Topology
	NavigateTo "http://localhost:32843/Topology/topology.svc"
	
	# Warm up Host Name Site Collections (HNSC)
	Write-Output "Opening Host Name Site Collections (HNSC)..."
	$hnsc = Get-SPSite -Limit All | Where-Object {$_.HostHeaderIsSiteName -eq $true} | Select-Object Url
	foreach ($sc in $hnsc) {
		NavigateTo $sc.Url
	}
}

Function NavigateTo([string] $url) {
	if ($url.ToUpper().StartsWith("HTTP")) {
		Write-Host "  $url" -NoNewLine
		# WebRequest command line
		try {
			$wr = Invoke-WebRequest -Uri $url -UseBasicParsing -UseDefaultCredentials -TimeoutSec 120
			FetchResources $url $wr.Images
			FetchResources $url $wr.Scripts
			Write-Host "."
		} catch {
			$httpCode = $_.Exception.Response.StatusCode.Value__
			if ($httpCode) {
				Write-Host "   [$httpCode]" -Fore Yellow
			} else {
				Write-Host " "
			}
		}
	}
}

Function FetchResources($baseUrl, $resources) {
	# Download additional HTTP files
	[uri]$uri = $baseUrl
	$rootUrl = $uri.Scheme + "://" + $uri.Authority
	
	# Loop
	$counter = 0
	foreach ($res in $resources) {
		# Support both abosolute and relative URLs
		$resUrl  = $res.src
		if ($resUrl.ToUpper().Contains("HTTP")) {
			$fetchUrl = $res.src
		} else {
			if (!$resUrl.StartsWith("/")) {
				$resUrl = "/" + $resUrl
			}
			$fetchUrl = $rootUrl + $resUrl
		}

		# Progress
		Write-Progress -Activity "Opening " -Status $fetchUrl -PercentComplete (($counter/$resources.Count)*100)
		$counter++
		
		# Execute
		$resp = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $fetchUrl -TimeoutSec 120
		Write-Host "." -NoNewLine
	}
	Write-Progress -Activity "Completed" -Completed
}

Function ShowW3WP() {
	# Total memory used by IIS worker processes
	$mb = [Math]::Round((Get-Process W3WP -ErrorAction SilentlyContinue | Select-Object pm | Measure-Object pm -Sum).Sum/1MB)
	WriteLog "Total W3WP = $mb MB" "Green"
}

Function CreateLog() {
	# EventLog - create source if missing
	if (!(Get-EventLog -LogName Application -Source "SPBestWarmUp" -ErrorAction SilentlyContinue)) {
		New-EventLog -LogName Application -Source "SPBestWarmUp" -ErrorAction SilentlyContinue | Out-Null
	}
}

Function WriteLog($text, $color) {
	$global:msg += "`n$text"
	if ($color) {
		Write-Host $text -Fore $color
	} else {
		Write-Output $text
	}
}

Function SaveLog($id, $txt, $error) {
	# EventLog
	if (!$skiplog) {
		if (!$error) {
			# Success
			$global:msg += $txt
			Write-EventLog -LogName Application -Source "SPBestWarmUp" -EntryType Information -EventId $id -Message $global:msg
		} else {      
			# Error
			$global:msg += $error[0].Exception.ToString() + "`r`n" + $error[0].ErrorDetails.Message
			Write-EventLog -LogName Application -Source "SPBestWarmUp" -EntryType Warning -EventId $id -Message $global:msg
		}
	}
}

# Main
CreateLog
WriteLog "SPBestWarmUp v2.2.4  (last updated 05-13-2016)`n------`n"

# Check Permission Level
if (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
	Write-Warning "You do not have elevated Administrator rights to run this script.`nPlease re-run as Administrator."
	break
} else {
	try {
		# SharePoint cmdlets
		Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null

		# Task Scheduler
		$cmdpath = $MyInvocation.MyCommand.Path
		$tasks = schtasks /query /fo csv | ConvertFrom-Csv
		$spb = $tasks |Where-Object {$_.TaskName -eq "\SPBestWarmUp"}
		if (!$spb -and !$install -and !$installfarm) {
			Write-Warning "Tip: to install on Task Scheduler run the command ""SPBestWarmUp.ps1 -install"""
		}
		if ($install -or $installfarm -or $uninstall) {
			Installer
			SaveLog 2 "Installed to Task Scheduler"
		}
		if ($uninstall) {
			break
		}
		
		# Core
		ShowW3WP
		WarmUp
		ShowW3WP
		
		# Custom URLs - Add your own below
		# Looks at Central Admin Site Title to support many farms with a single script
		(Get-SPWebApplication -IncludeCentralAdministration) |Where-Object {$_.IsAdministrationWebApplication -eq $true} |ForEach-Object {
			$caTitle = Get-SPWeb $_.Url | Select-Object Title
		}
		switch -Wildcard ($caTitle) {
			"*PROD*" {
				#NavigateTo "http://portal/popularPage.aspx"
				#NavigateTo "http://portal/popularPage2.aspx"
				#NavigateTo "http://portal/popularPage3.aspx
			}
			"*TEST*" {
				#NavigateTo "http://portal/popularPage.aspx"
				#NavigateTo "http://portal/popularPage2.aspx"
				#NavigateTo "http://portal/popularPage3.aspx
			}
			default {
				#NavigateTo "http://portal/popularPage.aspx"
				#NavigateTo "http://portal/popularPage2.aspx"
				#NavigateTo "http://portal/popularPage3.aspx
			}
		}
		SaveLog 1 "Operation completed successfully"
	} catch {
		SaveLog 201 "ERROR" $error
	}
}
