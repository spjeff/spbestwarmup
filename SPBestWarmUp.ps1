<#
.SYNOPSIS
	Warm up SharePoint IIS W3WP memory cache by loading pages from WebRequest

.DESCRIPTION
	Loads the full page so resources like CSS, JS, and images are included. Please modify lines 374-395 to suit your portal content design (popular URLs, custom pages, etc.)
	
	Comments and suggestions always welcome!  Please, use the issues panel at the project page.

.PARAMETER url
	A collection of url that will be added to the list of websites the script will fetch.
	
.PARAMETER install
	Typing "SPBestWarmUp.ps1 -install" will create a local Task Scheduler job under credentials of the current user. Job runs every 60 minutes on the hour to help automatically populate cache. Keeps cache full even after IIS daily recycle, WSP deployment, reboot, or other system events.

.PARAMETER installfarm
	Typing "SPBestWarmUp.ps1 -installfarm" will create a Task Scheduler job on all machines in the farm.

.PARAMETER uninstall
	Typing "SPBestWarmUp.ps1 -uninstall" will remove Task Scheduler job from all machines in the farm.
	
.PARAMETER user
	Typing "SPBestWarmUp.ps1 -user" provides the user name that will be used for the execution of the Task Scheduler job. If this parameter is missing it is assumed that the Task Scheduler job will be run with the current user.
	
.PARAMETER skiplog
	Typing "SPBestWarmUp.ps1 -skiplog" will avoid writing to the EventLog.
	
.PARAMETER allsites
	Typing "SPBestWarmUp.ps1 -allsites" will load every site and web URL. If the parameter skipsubwebs is used, only the root web of each site collection will be processed.

.PARAMETER skipsubwebs
	Typing "SPBestWarmUp.ps1 -skipsubwebs" will skip the subwebs of each site collection and only process the root web of the site collection.

.PARAMETER skipadmincheck
	Typing "SPBestWarmUp.ps1 -skipadmincheck" will skip checking of the current user is a local administrator. Local administrator rights are necessary for the installation of the Windows Task Scheduler but not necessary for simply running the warmup script.

.EXAMPLE
	.\SPBestWarmUp.ps1 -url "http://domainA.tld","http://domainB.tld"

.EXAMPLE
	.\SPBestWarmUp.ps1 -i
	.\SPBestWarmUp.ps1 -install

.EXAMPLE
	.\SPBestWarmUp.ps1 -f
	.\SPBestWarmUp.ps1 -installfarm

.EXAMPLE
	.\SPBestWarmUp.ps1 -f -user "Contoso\JaneDoe"
	.\SPBestWarmUp.ps1 -installfarm -user "Contoso\JaneDoe"

.EXAMPLE
	.\SPBestWarmUp.ps1 -u
	.\SPBestWarmUp.ps1 -uninstall

	
.NOTES  
	File Name:  SPBestWarmUp.ps1
	Author   :  Jeff Jones  - @spjeff
	Author   :  Hagen Deike - @hd_ka
	Author   :  Lars Fernhomberg
	Author   :  Charles Crossan - @crossan007
	Author   :  Leon Lennaerts - SPLeon
	Version  :  2.4.18
	Modified :  2018-02-20

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
	
	[Parameter(Mandatory=$False, Position=4, ValueFromPipeline=$false, HelpMessage='Use -user to provide the login of the user that will be used to run the script in the Windows Task Scheduler job')]
	[string]$user,
	
	[Parameter(Mandatory=$False, Position=5, ValueFromPipeline=$false, HelpMessage='Use -skiplog -sl parameter to avoid writing to Event Log')]
	[Alias("sl")]
	[switch]$skiplog,
	
	[Parameter(Mandatory=$False, Position=6, ValueFromPipeline=$false, HelpMessage='Use -allsites -all parameter to load every site and web (if skipsubwebs parameter is also given, only the root web will be processed)')]
	[Alias("all")]
	[switch]$allsites,

	[Parameter(Mandatory=$False, Position=7, ValueFromPipeline=$false, HelpMessage='Use -skipsubwebs -sw parameter to skip subwebs of each site collection and to process only the root web')]
	[Alias("sw")]
	[switch]$skipsubwebs,

	[Parameter(Mandatory=$False, Position=8, ValueFromPipeline=$false, HelpMessage='Use -skipadmincheck -sac parameter to skip checking if the current user is an administrator')]
	[Alias("sac")]
	[switch]$skipadmincheck,

	[Parameter(Mandatory=$False, Position=9, ValueFromPipeline=$false, HelpMessage='Use -skipserviceapps -ssa parameter to skip warming up of Service Application Endpoints URLs')]
	[Alias("ssa")]
	[switch]$skipserviceapps,

	[Parameter(Mandatory=$False, Position=10, ValueFromPipeline=$false, HelpMessage='Use -skipprogress -sp parameter to skip display of progres bar.  Faster execution for background scheduling.')]
	[Alias("sp")]
	[switch]$skipprogress
)

Function Installer() {
	# Add to Task Scheduler
	Write-Output "  Installing to Task Scheduler..."
	if(!$user) {
		$user = $ENV:USERDOMAIN + "\"+$ENV:USERNAME
	}
	Write-Output "  User for Task Scheduler job: $user"
	
    # Attempt to detect password from IIS Pool (if current user is local admin and farm account)
    $appPools = Get-WMIObject -Namespace "root/MicrosoftIISv2" -Class "IIsApplicationPoolSetting" | Select-Object WAMUserName, WAMUserPass
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
	$suffix += " -skipadmincheck"	#We do not need administrative rights on local machines to check the farm
	if ($allsites) {$suffix += " -allsites"}
	if ($skipsubwebs) {$suffix += " -skipsubwebs"}
	if ($skiplog) {$suffix += " -skiplog"}
	if ($skipprogress) {$suffix += " -skipprogress"}
	$cmd = "-ExecutionPolicy Bypass -File SPBestWarmUp.ps1" + $suffix
	
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
			WriteLog "  [OK]" Green
		} else {
			$xmlCmdPath = $cmdpath.Replace(".ps1", ".xml")
			# Ensure that XML file is present
			if(!(Test-Path $xmlCmdPath)) {
				Write-Warning """$($xmlCmdPath)"" is missing. Cannot create timer job without missing file."
				return
			}

			# Update xml file
			Write-Host "xmlCmdPath - $xmlCmdPath"
			$xml = [xml](Get-Content $xmlCmdPath)
			$xml.Task.Principals.Principal.UserId = $user
			$xml.Task.Actions.Exec.Arguments = $cmd
			$xml.Task.Actions.Exec.WorkingDirectory = (Split-Path ($xmlCmdPath)).ToString()
			$xml.Save($xmlCmdPath)

			# Copy local file to remote UNC path machine
			Write-Output "SCHTASKS CREATE on $_"
			if ($_ -ne "localhost" -and $_ -ne $ENV:COMPUTERNAME) {
				$dest = $cmdpath
				$drive = $dest.substring(0,1)
				$match =  Get-WMIObject -Class Win32_LogicalDisk | Where-Object {$_.DeviceID -eq ($drive+":") -and $_.DriveType -eq 3}
				if ($match) {
					$dest = "\\" + $_ + "\" + $drive + "$" + $dest.substring(2,$dest.length-2)
					$xmlDest = $dest.Replace(".ps1", ".xml")
					mkdir (Split-Path $dest) -ErrorAction SilentlyContinue | Out-Null
					Write-Output $dest
					Copy-Item $cmdpath $dest -Confirm:$false
					Copy-Item $xmlCmdPath $xmlDest -Confirm:$false
				}
			}
			# Create task
			schtasks /s $_ /create /tn "SPBestWarmUp" /ru $user /rp $pass /xml $xmlCmdPath
			WriteLog "  [OK]"  Green
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

	# Accessing the Alternate URls to warm up all "extended webs" (i.e. multiple IIS websites exists for one SharePoint webapp)
	$was = Get-SPWebApplication -IncludeCentralAdministration
	foreach ($wa in $was) {
		foreach ($alt in $wa.AlternateUrls) {
			$url = $alt.PublicUrl
			if(!$url.EndsWith("/")) {
				$url = $url + "/"
			}
			NavigateTo $url
			NavigateTo $url"_api/web"
			NavigateTo $url"_api/_trust" # for ADFS, first user login
			NavigateTo $url"_layouts/viewlsts.aspx"
			NavigateTo $url"_layouts/settings.aspx"
			NavigateTo $url"_vti_bin/UserProfileService.asmx"
			NavigateTo $url"_vti_bin/sts/spsecuritytokenservice.svc"
			NavigateTo $url"_api/search/query?querytext='warmup'"

			# SharePoint 2016
			NavigateTo $url"_layouts/15/fonts/shellicons.eot"
			NavigateTo $url"_layouts/15/jsgrid.js"
			NavigateTo $url"_layouts/15/sp.js"
			NavigateTo $url"_layouts/15/sp.ribbon.js"
			NavigateTo $url"_layouts/15/core.js"
			NavigateTo $url"_layouts/15/init.js"
			NavigateTo $url"_layouts/15/cui.js"
			NavigateTo $url"_layouts/15/inplview.js"
			NavigateTo $url"_layouts/15/suitenav.js"
		}
		
		# Warm Up Individual Site Collections and Sites
		if ($allsites) {
 			$sites = (Get-SPSite -WebApplication $wa -Limit ALL)
 			foreach($site in $sites) {
				if($skipsubwebs)
				{
					$url = $site.RootWeb.Url
					NavigateTo $url
				}
				else
				{
					$webs = (Get-SPWeb -Site $site -Limit ALL)
					foreach($web in $webs){
						$url = $web.Url
						NavigateTo $url
					}
				}
 			}
		}
		
        # Central Admin
        if ($wa.IsAdministrationWebApplication) {
			$url = $wa.Url
			# Specific pages
            NavigateTo $url"Lists/HealthReports/AllItems.aspx"
            NavigateTo $url"_admin/FarmServers.aspx"
            NavigateTo $url"_admin/Server.aspx"
            NavigateTo $url"_admin/WebApplicationList.aspx"
			NavigateTo $url"_admin/ServiceApplications.aspx"
			
			# Quick launch top links
			NavigateTo $url"applications.aspx"
			NavigateTo $url"systemsettings.aspx"
			NavigateTo $url"monitoring.aspx"
			NavigateTo $url"backups.aspx"
			NavigateTo $url"security.aspx"
			NavigateTo $url"security.aspx"
			NavigateTo $url"upgradeandmigration.aspx"
			NavigateTo $url"apps.aspx"
			NavigateTo $url"office365configuration.aspx"
			NavigateTo $url"generalapplicationsettings.aspx"

            # Manage Service Application
            $sa = Get-SPServiceApplication
            $links = $sa | ForEach-Object {$_.ManageLink.Url} | Select-Object -Unique
            foreach ($link in $links) {
                $ml = $link.TrimStart('/')
                NavigateTo "$url$ml"
            }
        }
    }
	
    # Warm up Service Applications
	if (!$skipserviceapps) {
    	Get-SPServiceApplication | ForEach-Object {$_.EndPoints | ForEach-Object {$_.ListenUris | ForEach-Object {NavigateTo $_.AbsoluteUri}}}
	}

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

	# Warm up Office Online Server (OOS)
	$remoteuis = "m,o,oh,op,p,we,wv,x".Split(",")
	$services = "diskcache/DiskCache.svc,dss/DocumentSessionService.svc,ecs/ExcelService.asmx,farmstatemanager/FarmStateManager.svc,metb/BroadcastStateService.svc,pptc/Viewing.svc,ppte/Editing.svch,wdss/WordDocumentSessionService.svc,wess/WordSaveService.svc,wvc/Conversion.svc".Split(",")

	# Loop per WOPI
	$wopis = Get-SPWOPIBinding | Select-Object ServerName -Unique
	foreach ($w in $wopis.ServerName) {
		foreach ($r in $remoteuis) {
			NavigateTo "http://$w/$r/RemoteUIs.ashx"
			NavigateTo "https://$w/$r/RemoteUIs.ashx"
		}
		foreach ($s in $services) {
			NavigateTo "http://$w"+":809/$s/"
			NavigateTo "https://$w"+":810/$s/"
		}
	}
}

Function NavigateTo([string] $url) {
	if ($url.ToUpper().StartsWith("HTTP") -and !$url.EndsWith("/ProfileService.svc","CurrentCultureIgnoreCase")) {
		WriteLog "  $url" -NoNewLine
		# WebRequest command line
		try {
			$wr = Invoke-WebRequest -Uri $url -UseBasicParsing -UseDefaultCredentials -TimeoutSec 120
			FetchResources $url $wr.Images
			FetchResources $url $wr.Scripts
			Write-Host "."
		} catch {
			$httpCode = $_.Exception.Response.StatusCode.Value__
			if ($httpCode) {
				WriteLog "   [$httpCode]" Yellow
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
		if (!$skipprogress) {
        	Write-Progress -Activity "Opening " -Status $fetchUrl -PercentComplete (($counter/$resources.Count)*100)
			$counter++
		}
		
        # Execute
        Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $fetchUrl -TimeoutSec 120 | Out-Null
        Write-Host "." -NoNewLine
	}
	if (!$skipprogress) {
		Write-Progress -Activity "Completed" -Completed
	}
}

Function ShowW3WP() {
    # Total memory used by IIS worker processes
    $mb = [Math]::Round((Get-Process W3WP -ErrorAction SilentlyContinue | Select-Object workingset64 | Measure-Object workingset64 -Sum).Sum/1MB)
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
			$global:msg += "ERROR`n"
            $global:msg += $error.Message + "`n" + $error.ItemName
            Write-EventLog -LogName Application -Source "SPBestWarmUp" -EntryType Warning -EventId $id -Message $global:msg
        }
    }
}

# Main
CreateLog
$cmdpath = (Resolve-Path .\).Path
$cmdpath += "\SPBestWarmUp.ps1"
$ver = $PSVersionTable.PSVersion
WriteLog "SPBestWarmUp v2.4.18  (last updated 2018-02-20)`n------`n"
WriteLog "Path: $cmdpath"
WriteLog "PowerShell Version: $ver"

# Check Permission Level
if (!$skipadmincheck -and !([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
	Write-Warning "You do not have elevated Administrator rights to run this script.`nPlease re-run as Administrator."
	break
} else {
    try {
        # SharePoint cmdlets
        Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null

        # Task Scheduler
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
		SaveLog 101 "ERROR" $_.Exception
	}
}
