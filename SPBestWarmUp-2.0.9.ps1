#requires -version 3.0
<#  
.SYNOPSIS  
    Warm up SharePoint IIS memory cache by touching pages with Invoke-WebRequest
.DESCRIPTION  
    The Warm-Up script prefetches SharePoints ASPX pages and loads them into the IIS cache. This will help
    to improve the user experience. The script normaly doesn't load pictures and javascript files, but can
    be configured to load the static content as well to preload caches.
    Comments and suggestions are always welcome! Please, use the dicussions or issues panel at the project page.
.PARAMETER Url
	A collection of url that will be added to the list of websites the script will fetch.
.PARAMETER Install
	Typing "SPBestWarmUp.ps1 -install" will create a local Task Scheduler job under credentials of the
	current user.  Job runs every 60 minutes on the hour to help automatically populate cache after
	nightly IIS recycle.
.PARAMETER Transcript
	Creates a transcript log on the disk for later inspection.
.PARAMETER FetchStaticContent
    Without this switch enabled, the script will only fetch the HTML page from the SharePoint server. The static
    content, like e.g. pictures won't be pulled. By using this parameter the script will fetch all the
    pictures and javascripts from the server as well. This can be used to warm up the static caches on the
    servers or proxies.
.EXAMPLE
    .\SPBestWarmUp.ps1 -Install
    Installs the script in the Windows scheduler
.EXAMPLE
    .\SPBestWarmUp.ps1 -Transcript c:\Temp\log.txt
    Creates the translog file at the given path
.EXAMPLE
    .\SPBestWarmUp.ps1 -Url "http://domainA.tld","https://domainB.tld"
    Adds the given urls to the list that wil be fetched when the script is executed.
.NOTES  
    File Name     : SPBestWarmUp.ps1
    Author        : Jeff Jones - @spjeff
                  : Hagen Deike - @hd_ka
    Version       : 2.0.9 (Stockholm)
	Last Modified : 11-09-2015
.LINK
	https://spbestwarmup.codeplex.com/documentation
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$False, Position=0, ValueFromPipeline=$false, HelpMessage='A collection of Url that will be fetched too')]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("https?:\/\/\D+")]
    [string[]]$Url,

	[Parameter(Mandatory=$False, Position=1, ValueFromPipeline=$false, HelpMessage='Use the Install parameter to install the script to the Windows scheduler')]
	[Alias("i")]
    [switch]$Install,

    [Parameter(Mandatory=$False, Position=2, ValueFromPipeline=$false, HelpMessage='Fetch static content from SharePoint too')]
    [switch]$FetchStaticContent,

	[Parameter(Mandatory=$False, Position=3, ValueFromPipeline=$false, HelpMessage='Define the path where the transcript should be logged')]
    [ValidateNotNullOrEmpty()]
	[string]$Transcript
)

Function Installer {
# Add to Task Scheduler

	WriteScreenLog -Message "Installing to Task Scheduler..." -printTime
	$user = $ENV:USERDOMAIN+"\"+$ENV:USERNAME
	WriteScreenLog -Message "Current User: $user" -printTime
	
	# Attempt to detect password from IIS Pool (if current user is local admin and farm account)
	$appPools = Get-WmiObject -Namespace "root\MicrosoftIISV2" -Class "IIsApplicationPoolSetting" | Select WAMUserName, WAMUserPass
	Write-Verbose -Message ""Found $appPools.Count Application Pool""

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
	$cmd = """PowerShell.exe -ExecutionPolicy Bypass '$global:path'"""
	schtasks /create /tn "SPBestWarmUp" /ru $user /rp $pass /rl highest /sc daily /st 01:00 /ri 60 /du 24:00 /tr $cmd
	WriteScreenLog -Message "Task created" -Type OK -printTime
	Write-Host
}
Function GetWebApplicationUrls {
	$webApplications = Get-SPWebApplication -IncludeCentralAdministration
	$webApplications |? {$_.IsAdministrationWebApplication -eq $true} |% {$caTitle = Get-SPWeb $_.Url | Select Title}

	# Warm up SharePoint web applications
	WriteScreenLog -Message "Collecting Web Applications..." -printTime

	foreach ($webApplication in $webApplications) {
        "Processing web application {0}" -f $webApplication.Url | Write-Verbose
		$global:siteUrls += $webApplication.Url
		$global:siteUrls += $webApplication.Url + "_layouts/viewlsts.aspx"
		$global:siteUrls += $webApplication.Url + "_vti_bin/UserProfileService.asmx"
		$global:siteUrls += $webApplication.Url + "_vti_bin/sts/spsecuritytokenservice.svc"
	}
}
Function GetServiceApplicationUrls {
	# Warm up Service Applications
	$serviceApplications = Get-SPServiceApplication
    WriteScreenLog -Message "Collecting Service Applications..." -printTime

	foreach($serviceApplication in $serviceApplications) {

        "Processing service application {0} Id: {1} " -f $serviceApplication.DisplayName, $serviceApplication.id | Write-Verbose
		foreach($endpoint in $serviceApplication.Endpoints) {
            foreach($listenUri in $endpoint.ListenUris) {
				# Remove all endpoints not matching http[s]
				if($listenUri.AbsoluteUri -match "https?:\/\/") {
					$global:siteUrls += $listenUri.AbsoluteUri -replace "\/https?", ""
				}
			}
		}
	}
}
Function GetHostNameSiteCollectionsUrls {
	# Warm up Host Name Site Collections (HNSC)
	WriteScreenLog -Message "Collecting Host Name Site Collections (HNSC)..." -printTime
	$hnsc = Get-SPSite -Limit All |? {$_.HostHeaderIsSiteName -eq $true} | Select Url

	foreach ($sc in $hnsc) {
        "Adding {0} to the collection of URLs to fetch" -f $sc.Url | Write-Verbose
		$global:siteUrls += $sc.Url
	}
}
Function GetExternalUrls {
	WriteScreenLog -Message "Collecting external urls..." -printTime
    foreach($extUrl in $Url) {
        "Adding {0} to the collection of URLs to fetch" -f $extUrl | Write-Verbose
        $global:siteUrls += $extUrl
    }
}
Function NavigateTo {
    Param (
        [Parameter(Mandatory=$True,Position=0)]
        [string] $url
    )

	WriteScreenLog -Message "Opening $url" -Type Info -printTime
	try {
        $webReturn = Invoke-WebRequest -UseDefaultCredentials -Uri $url -TimeoutSec 120
	    if($webReturn.StatusDescription -eq "OK") {
            WriteScreenLog -Message $webReturn.StatusCode -Type Info -printTime		    
            WriteScreenLog -Message "Success..." -Type OK -printTime
	    }
        if($FetchStaticContent -eq $True) {
            
            $imageCounter = 0
            $Images = $webReturn.Images | select src -Unique
            foreach($image in $Images) {
                $imageUrl = $url + $image.src
                WriteScreenLog -Message "Opening $imageUrl" -Type Info -printTime
                Write-Progress -Activity "Fetching images" -status $imageUrl -Id 2 -ParentId 1 -PercentComplete (($imageCounter/$Images.Count)*100)

                $imageReturn = Invoke-WebRequest -UseDefaultCredentials -Uri $imageUrl -TimeoutSec 120
                
                $imageCounter++
	            if($imageReturn.StatusDescription -eq "OK") {
                    WriteScreenLog -Message $webReturn.StatusCode -Type Info -printTime		    
                    WriteScreenLog -Message "Success..." -Type OK -printTime
	            }
            }

            $scriptCounter = 0
            $Scripts = $webReturn.Scripts | select src -Unique
            foreach($script in $Scripts) {
                $scriptUrl = $url + $script.src
                WriteScreenLog -Message "Opening $scriptUrl" -Type Info -printTime
                Write-Progress -Activity "Fetching scripts" -status $scriptUrl -Id 2 -ParentId 1 -PercentComplete (($scriptCounter/$Scripts.Count)*100)
                
                $scriptReturn = Invoke-WebRequest -UseDefaultCredentials -Uri $scriptUrl -TimeoutSec 120
                
                $scriptCounter++
        	    if($scriptReturn.StatusDescription -eq "OK") {
                    WriteScreenLog -Message $webReturn.StatusCode -Type Info -printTime		    
                    WriteScreenLog -Message "Success..." -Type OK -printTime
	            }
            }

            Write-Progress -Activity "Done..." -Id 2 -ParentId 1 -PercentComplete 100 -Completed
        }
    }
    catch {
        WriteScreenLog -Message $_.Exception.Message -Type Warning -printTime		    
        WriteScreenLog -Message "Url cound not be opened" -Type Warning -printTime

        $message = $url
        $message += "`r`n" # CR+LF
        #$message += $_.Exception.Message
        $message += $error[0].Exception
        $message += "`r`n" # CR+LF
        $message += $error[0].ErrorDetails.Message
        Write-EventLog -LogName Application -Source "SPBestWarmUp" -EntryType Warning -EventId 201 -Message $message
    }
}
function WriteScreenLog {
##
## Write a log entry to the screen
## Usage:
##        WriteScreenLog -Message "Message text" [-Type OK|Warning|Error|Info|Verbose] [-printTime]

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[string]$Message,
		
		[ValidateSet("OK","Warning","Error", "Info", "Verbose")] 
		[string]$Type,
		
		[switch]$printTime
    )
	$screenXpos = [Math]::Truncate($Host.UI.RawUI.WindowSize.Width - 11)

	# Write the message to the screen
	$now = ""
	if($printTime -eq $true){
		$now = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
		$now = "$now | "
	}
    $Message = $now + $Message
    Write-Output $Message
	
	if($Type -ne "") {
        [Console]::SetCursorPosition($screenXpos, $Host.UI.RawUI.CursorPosition.Y-1)
	}
    switch ($Type) {
		"OK" {Write-Host -BackgroundColor Green -ForegroundColor Black  "    OK    "}
		"Warning" {Write-Host -BackgroundColor Yellow -ForegroundColor Black "  Warning "}
		"Error" {Write-Host -ForegroundColor Yellow -BackgroundColor Red  "   Error  "}
		"Info" {Write-Host -BackgroundColor $Host.UI.RawUI.ForegroundColor -ForegroundColor $Host.UI.RawUI.BackgroundColor "   Info   "}
		"Verbose" {Write-Host -BackgroundColor $Host.UI.RawUI.ForegroundColor -ForegroundColor $Host.UI.RawUI.BackgroundColor "  Verbose "}
	}
}
Function WarmUp {
##
## The Main WarmUp function
##
		
	# Load the SharePoint PowerShell cmdlets
    "Testing for the Microsoft.SharePoint.PowerShell module" | Write-Verbose
	if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ea silentlycontinue) -eq $null)
	{
		WriteScreenLog -Message "Loading the SharePoint PowerShell module" -printTime
		Add-PSSNapin Microsoft.SharePoint.Powershell
	}

	$global:siteUrls = @()
	GetWebApplicationUrls
	GetServiceApplicationUrls
	GetHostNameSiteCollectionsUrls
    "Adding topology.svc to the collection of url to fetch" | Write-Verbose
	$global:siteUrls += "http://localhost:32843/Topology/topology.svc"
    GetExternalUrls

    $progressCounter = 0
	foreach($target in $global:siteUrls) {
        Write-Progress -Activity "Fetching webpages from SharePoint server" -status $target -Id 1 -PercentComplete (($progressCounter/$Global:siteUrls.Count)*100)
        NavigateTo -url $target
        Write-Progress -Activity "Done..." -Id 1 -PercentComplete 100 -Completed
        $progressCounter++
	}
}
Function CollectSystemInformation {
    Process {
        "PowerShell Version: {0}" -f $Host.Version | Write-Verbose
        "PowerShell Culture: {0}" -f $Host.CurrentCulture | Write-Verbose
        "PowerShell ExecutionPolicy: {0}" -f (Get-ExecutionPolicy) | Write-Verbose
		gwmi win32_operatingsystem | % caption
    }
}

##
## Main
##
if ($PSBoundParameters['Verbose']) {
    CollectSystemInformation
}

# Load the EventLog
if ((Get-EventLog -LogName Application -Source "SPBestWarmUp" -ea silentlycontinue) -eq $null)
{
    New-EventLog -LogName Application -Source "SPBestWarmUp"
	WriteScreenLog -Message "Windows EventLog for SPBestWarmUp created..." -printTime
}

#if ($Transcript -and (Test-Path -Path $Transcript -PathType Container) -replace "\\\w+.\w+\Z","") {
if ($Transcript) {
    Write-Verbose -Message "Starting Transscript"
    Start-Transcript -Path $Transcript -Force # See issue 22023
}
WriteScreenLog -Message "SPBestWarmUp started..." -printTime

# Check for permission level. If not sufficient, get elevated rights for the shell
Write-Verbose -Message "Checking for elevated rights"
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
	WriteScreenLog -Message "Reloading the PowerShell with elevated rights." -Type Info -printTime
	Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Write-Verbose -Message "PowerShell restarted with elevated rights"
	exit
}

If (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    WriteScreenLog -Message "You do not have Administration rights! Please, restart the PowerShell with Administrator rights" -Type Error -printTime
    break
} else {
    # Start the Warm up process
    $global:path = $MyInvocation.MyCommand.Path
    $tasks = schtasks /query /fo csv | ConvertFrom-Csv
    $spb = $tasks |? {$_.TaskName -eq "\SPBestWarmUp"}
    
	if (!$spb -and !$Install) {
	    WriteScreenLog -Message "To install on Task Scheduler run the command ""SPBestWarmUp.ps1 -install """ -Type Info -printTime
    }

    if ($Install) {
	    Installer
        Write-EventLog -LogName Application -Source "SPBestWarmUp" -EntryType Information -EventId 2 -Message "The script was installed to thw Windows scheduler successfully"
    } else {
        WarmUp
        Write-EventLog -LogName Application -Source "SPBestWarmUp" -EntryType Information -EventId 1 -Message "The script has run successfully"
    }
}

if ($Transcript) {
	Stop-Transcript
}
