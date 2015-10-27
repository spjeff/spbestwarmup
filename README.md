## Project Description
Tired of waiting for SharePoint pages to load? Want something easy to support? That works on all versions? This warmup script is for you!

I used several different warmup scripts over the years. They all worked OK but each seemed to lack one or two features so I decided to create one for myself. Hopefully you find it useful too.


## Get Started
* Excellent blog post by @hd_ka at http://blog.greenbrain.de/2014/10/fire-up-those-caches.html
* Grant permission
* Configure Task Scheduler
* Enable trigger conditions

## Key Features
* Supports both SharePoint 2010 and 2013
* Supports custom page URLs
* Automatically detects all Web Application URLs
* Downloads full page resources (CSS, JS, images) not just HTML
* Downloads using Internet Explorer COM automation
* Great for ECM websites to help populate blob cache
* Warms up Central Admin too. Faster admin UI experience!

## Quick Start
* Download the release, unpack and rename the script to "SPBestWarmup.ps1"
* Copy "SPBestWarmup.ps1" on each SharePoint web front end (WFE)
* Run "SPBestWarmup.ps1 -install" to create the Task Scheduler item
* Sit back and watch it run

## Admin Tip
* After reboot run this command to manually trigger the job and warm up IIS
SCHTASKS /RUN /TN "SPBestWarmup"

## Screenshots

* Run with Scheduled Task present
![image](https://raw.githubusercontent.com/spjeff/spbestwarmup/master/doc/1.png)
* Run without Scheduled Task (reminder how to create)
![image](https://raw.githubusercontent.com/spjeff/spbestwarmup/master/doc/2.png)
* Install to create Scheduled Task
![image](https://raw.githubusercontent.com/spjeff/spbestwarmup/master/doc/3.png)

## Note
* Running this with a different service account than farm might require you to first grant PowerShell access. This will ensure the service account has access to run "Get-SPWebApplication" and read ConfigDB for which URLs to load. http://technet.microsoft.com/en-us/magazine/gg490648.aspx

## Contact
Please drop a line to [@spjeff](https://twitter.com/spjeff) or [spjeff@spjeff.com](mailto:spjeff@spjeff.com)
Thanks!  =)
