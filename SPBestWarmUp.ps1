<#
.SYNOPSIS  
    Warm up SharePoint IIS W3WP memory cache by loading pages from Internet Explorer or WebRequest

.DESCRIPTION
    Loads the full page so resources like CSS, JS, and images are included.  Please modify lines 85-105
	to suit your portal content design (popular URLs, custom pages, etc.)
    
	Comments and suggestions always welcome!  Please, use the dicussions or issues panel at the project page.

.PARAMETER install
	Typing "SPBestWarmUp.ps1 -install" will create a local Task Scheduler job under credentials of the
	current user.  Job runs every 60 minutes on the hour to help automatically populate cache after nightly IIS recycle.

.EXAMPLE
	.\SPBestWarmUp.ps1 -url "http://domainA.tld","http://domainB.tld"

.EXAMPLE
    .\SPBestWarmUp.ps1 -i
	.\SPBestWarmUp.ps1 -install

.EXAMPLE
	.\SPBestWarmUp.ps1 -wr
    .\SPBestWarmUp.ps1 -webrequest
	
.NOTES  
    File Name     : SPBestWarmUp.ps1
    Author        : Jeff Jones  - @spjeff
					Hagen Deike - @hd_ka
    Version       : 2.11
	Last Modified : 03-20-2016

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

	[Parameter(Mandatory=$False, Position=1, ValueFromPipeline=$false, HelpMessage='Use -Install parameter to add script to the Windows Task Scheduler')]
	[Alias("i")]
    [switch]$install,

	[Parameter(Mandatory=$False, Position=2, ValueFromPipeline=$false, HelpMessage='Use -WebRequest parameter to load HTTP with Invoke-WebRequest instead of Internet Explorer GUI')]
	[Alias("wr")]
    [switch]$webrequest
)

Function Installer() {
	# Add to Task Scheduler
	Write-Output "  Installing to Task Scheduler..."
	$user = $ENV:USERDOMAIN+"\"+$ENV:USERNAME
	Write-Output "  Current User: $user"
	
	# Attempt to detect password from IIS Pool (if current user is local admin and farm account)
	$appPools = Get-WmiObject -Namespace "root\MicrosoftIISV2" -Class "IIsApplicationPoolSetting" | Select WAMUserName, WAMUserPass
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
	
	# Create Task
	if ($webrequest) {$global:path += " -wr"}
	$cmd = """PowerShell.exe -ExecutionPolicy Bypass '$global:path'"""
	schtasks /create /tn "SPBestWarmUp" /ru $user /rp $pass /rl highest /sc daily /st 01:00 /ri 60 /du 24:00 /tr $cmd
	Write-Output "  [OK]"
}

Function WarmUp() {
	# Load plugin
	Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
	$was = Get-SPWebApplication -IncludeCentralAdministration
	$was |? {$_.IsAdministrationWebApplication -eq $true} |% {$caTitle = Get-SPWeb $_.Url | Select Title}
	
	# Open Internet Explorer
	if (!$webrequest) {
		Write-Output "Opening Internet Explorer..."
		$global:ie = New-Object -Com "InternetExplorer.Application"
		$global:ie.Navigate("about:blank")
		$global:ie.Visible = $true
		$global:ieproc = (Get-Process -Name iexplore)|? {$_.MainWindowHandle -eq $global:ie.HWND}
	}
	
    # Warm up CMD parameter URLs
    $cmdurl |% {NavigateTo $_}

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
	Get-SPServiceApplication |% {$_.EndPoints |% {$_.ListenUris |% {NavigateTo $_.AbsoluteUri}}}

    # Warm up Project Server
	Write-Output "Opening Project Server PWAs..."
    if ((Get-Command Get-SPProjectWebInstance -ErrorAction SilentlyContinue).Count -gt 0) {
        Get-SPProjectWebInstance |% {
			# Thanks to Eugene Pavlikov for the snippet
			$url = ($_.Url).AbsoluteUri + "/"
		
			NavigateTo $url
			NavigateTo $url + "_layouts/viewlsts.aspx"
			NavigateTo $url + "_vti_bin/UserProfileService.asmx"
			NavigateTo $url + "_vti_bin/sts/spsecuritytokenservice.svc"
			NavigateTo $url + "Projects.aspx"
			NavigateTo $url + "Approvals.aspx"
			NavigateTo $url + "Tasks.aspx"
			NavigateTo $url + "Resources.aspx"
			NavigateTo $url + "ProjectBICenter/Pages/Default.aspx"
			NavigateTo $url + "_layouts/15/pwa/Admin/Admin.aspx"
		}
	}

	# Warm up Topology
	NavigateTo "http://localhost:32843/Topology/topology.svc"
	
	# Add your own URLs below.  Looks at Central Admin Site Title for full lifecycle support in a single script file.
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
		"*DEV*" {
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
	
	# Warm up Host Name Site Collections (HNSC)
	Write-Output "Opening Host Name Site Collections (HNSC)..."
	$hnsc = Get-SPSite -Limit All |? {$_.HostHeaderIsSiteName -eq $true} | Select Url
	foreach ($sc in $hnsc) {
		NavigateTo $sc.Url
	}
	
	# Cleanup
	if (!$webrequest) {
		# Close IE window
		if ($global:ie) {
			Write-Output "Closing IE"
			$global:ie.Quit()
		}
		$global:ieproc | Stop-Process -Force -ErrorAction SilentlyContinue
		
		# Clean Temporary Files
		Remove-item "$env:systemroot\system32\config\systemprofile\appdata\local\microsoft\Windows\temporary internet files\content.ie5\*.*" -Recurse -ErrorAction SilentlyContinue
		Remove-item "$env:systemroot\syswow64\config\systemprofile\appdata\local\microsoft\Windows\temporary internet files\content.ie5\*.*" -Recurse -ErrorAction SilentlyContinue
	}
}

Function NavigateTo([string] $url, [int] $delayTime = 500) {
    if ($url.ToUpper().StartsWith("HTTP")) {
        Write-Host "  $url" -NoNewLine
	    if ($webrequest) {
            # WebRequest command line    
            try {
				$wr = Invoke-WebRequest -Uri $url -UseBasicParsing -UseDefaultCredentials -TimeoutSec 120
				FetchResources $url "Images" $wr.Images
				FetchResources $url "Scripts" $wr.Scripts
				
            } catch {}
        } else {
		    # Internet Explorer
		    try {
			    $global:ie.Navigate($url)
		    } catch {
			    try {
				    $pid = $global:ieproc.id
			    } catch {}
			    Write-Output "  IE not responding.  Closing process ID $pid"
			    $global:ie.Quit()
			    $global:ieproc | Stop-Process -Force -ErrorAction SilentlyContinue
			    $global:ie = New-Object -Com "InternetExplorer.Application"
			    $global:ie.Navigate("about:blank")
			    $global:ie.Visible = $true
			    $global:ieproc = (Get-Process -Name iexplore)|? {$_.MainWindowHandle -eq $global:ie.HWND}
		    }
		    IEWaitForPage $delayTime
	    }
		Write-Host "."
    }
}

Function IEWaitForPage([int] $delayTime = 500) {
	# Wait for current page to finish loading
	$loaded = $false
	$loop = 0
	$maxLoop = 20
	while ($loaded -eq $false) {
		$loop++
		if ($loop -gt $maxLoop) {
			$loaded = $true
		}
		[System.Threading.Thread]::Sleep($delayTime) 
		# If the browser is not busy, the page is loaded
		if (-not $global:ie.Busy)
		{
			$loaded = $true
		}
	}
}

Function FetchResources($baseUrl, $type, $resources) {
	# Download additional HTTP files
	[uri]$uri = $baseUrl
	$rootUrl = $uri.Scheme + "://" + $uri.Authority
	
	# Loop
	$counter = 0
	foreach ($res in $resources) {
		# Support both abosolute and relative URLs
		$resUrl  = $res.src
		if ($resUrl -contains "HTTP") {
			$scriptUrl = $res.src
		} else {
			if (!$resUrl.StartsWith("/")) {
				$resUrl = "/" + $resUrl
			}
			$scriptUrl = $rootUrl + $resUrl
		}

		# Progress
		Write-Progress -Activity "Opening " -status $scriptUrl -Id 2 -ParentId 1 -PercentComplete (($counter/$resources.Count)*100)
		$counter++
		
		# Execute
		$scriptReturn = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $scriptUrl -TimeoutSec 120
		Write-Host "." -NoNewLine
	}
}

Function ShowW3WP() {
	$mb = [Math]::Round((Get-Process w3wp | Select pm | Measure pm -Sum).Sum/1MB)
	Write-Warning "Total W3WP = $mb MB"
}

# Main
Write-Output "SPBestWarmUp v2.11  (last updated 03-20-2016)`n------`n"

# Check Permission Level
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    break
} else {
    # Warm up
    $global:path = $MyInvocation.MyCommand.Path
    $tasks = schtasks /query /fo csv | ConvertFrom-Csv
    $spb = $tasks |? {$_.TaskName -eq "\SPBestWarmUp"}
    if (!$spb -and !$install) {
	    Write-Warning "Tip: to install on Task Scheduler run the command ""SPBestWarmUp.ps1 -install"""
    }
    if ($install) {
	    Installer
    }
	ShowW3WP
    WarmUp
	ShowW3WP
}