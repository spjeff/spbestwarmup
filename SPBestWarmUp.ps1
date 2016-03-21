<#
.SYNOPSIS  
    Warm up SharePoint IIS W3WP memory cache by loading pages from Internet Explorer or WebRequest

.DESCRIPTION
    Loads the full page so resources like CSS, JS, and images are included.  Please modify lines 85-105
	to suit your portal content design (popular URLs, custom pages, etc.)
    
	Comments and suggestions always welcome!  Please, use the dicussions or issues panel at the project page.

.PARAMETER url
	Typing "SPBestWarmUp.ps1 -install" will create a local Task Scheduler job under credentials of the
	current user.  Job runs every 60 minutes on the hour to help automatically populate cache.  
	Keeps cache full even after IIS daily recycle, WSP deployment, reboot, or other system events.
	
.PARAMETER install
	Typing "SPBestWarmUp.ps1 -install" will create a local Task Scheduler job under credentials of the
	current user.  Job runs every 60 minutes on the hour to help automatically populate cache.  
	Keeps cache full even after IIS daily recycle, WSP deployment, reboot, or other system events.
	
.EXAMPLE
	.\SPBestWarmUp.ps1 -url "http://domainA.tld","http://domainB.tld"

.EXAMPLE
    .\SPBestWarmUp.ps1 -i
	.\SPBestWarmUp.ps1 -install

.EXAMPLE
    .\SPBestWarmUp.ps1 -f
	.\SPBestWarmUp.ps1 -installfarm

	
.NOTES  
    File Name     : SPBestWarmUp.ps1
    Author        : Jeff Jones  - @spjeff
					Hagen Deike - @hd_ka
    Version       : 2.2
	Last Modified : 03-21-2016

.LINK
	http://spbestwarmup.codeplex.com/
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
    [switch]$uninstall
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
	
	# Command
	$cmd = """PowerShell.exe -ExecutionPolicy Bypass '$cmdpath -webrequest'"""
	
	# Target machines
	$machines = @()
	if ($installfarm -or $uninstall) {
		# Create farm wide on remote machines
		foreach ($srv in (Get-SPServer |Where-Object {$_.Role -ne "Invalid"})) {
			$machines += $srv.Address
		}
	} else {
		# Create local on current machine
		$machines += "localhost"
	}
	$machines |ForEach-Object {
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
				$match =  Get-CimInstance -Class Win32_LogicalDisk |Where-Object {$_.DeviceID -eq ($drive+":") -and $_.DriveType -ne 4}
				if ($match) {
					$dest = "\\" + $_ + "\" + $dest.substring(0,1) + "$" + $dest.substring(2,$dest.length-2)
					mkdir (Split-Path $dest) -ErrorAction SilentlyContinue
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
    $cmdurl |ForEach-Object {NavigateTo $_}

    # Warm up SharePoint web applications
	Write-Output "Opening Web Applications..."
	foreach ($wa in $was) {
		$url = $wa.Url
		NavigateTo $url
        NavigateTo $url"_api/web"
		NavigateTo $url"_layouts/viewlsts.aspx"
		NavigateTo $url"_vti_bin/UserProfileService.asmx"
		NavigateTo $url"_vti_bin/sts/spsecuritytokenservice.svc"
	}
	
	# Warm up Service Applications
	Get-SPServiceApplication |ForEach-Object {$_.EndPoints |ForEach-Object {$_.ListenUris |ForEach-Object {NavigateTo $_.AbsoluteUri}}}

    # Warm up Project Server
	Write-Output "Opening Project Server PWAs..."
    if ((Get-Command Get-SPProjectWebInstance -ErrorAction SilentlyContinue).Count -gt 0) {
        Get-SPProjectWebInstance |ForEach-Object {
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
	$hnsc = Get-SPSite -Limit All |Where-Object {$_.HostHeaderIsSiteName -eq $true} | Select-Object Url
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
			} catch {
				$httpCode = $_.Exception.Response.StatusCode.Value__
				Write-Error $httpCode
			}
		Write-Host "."
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
		if ($resUrl -contains "HTTP") {
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
		Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $fetchUrl -TimeoutSec 120
		Write-Host "." -NoNewLine
	}
}

Function ShowW3WP() {
	# Total memory used by IIS worker processes
	$mb = [Math]::Round((Get-Process W3WP -ErrorAction SilentlyContinue | Select-Object pm | Measure-Object pm -Sum).Sum/1MB)
	Write-Host "Total W3WP = $mb MB" -Fore Green
}

# Main
Write-Output "SPBestWarmUp v2.11  (last updated 03-21-2016)`n------`n"

# Check Permission Level
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    break
} else {
    # Task Scheduler
    $cmdpath = $MyInvocation.MyCommand.Path
    $tasks = schtasks /query /fo csv | ConvertFrom-Csv
    $spb = $tasks |Where-Object {$_.TaskName -eq "\SPBestWarmUp"}
    if (!$spb -and !$install) {
	    Write-Warning "Tip: to install on Task Scheduler run the command ""SPBestWarmUp.ps1 -install"""
    }
    if ($install -or $installfarm -or $uninstall) {
		Installer
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
}