## Description
Tired of waiting for SharePoint pages to load? Want something easy to support? That works on all versions? This warmup script is for you!

I used several different warmup scripts. All worked OK but each seemed to lack features so I decided to create one for myself. Hope you find it useful too.

## Key Features
* SharePoint 2010, 2013, and 2016
* Custom page URLs
* Detect SP Web Application URLs
* Detect Host Named Site Collection URLs
* Detects Service Application URLs
* Detects Project Server PWA 
* Download full page resources (JS, images) not just HTML
* Download with `Invoke-WebRequest` 
* Great for ECM websites to populate blob cache
* Warm up Central Admin for a faster admin GUI experience!
* Display W3WP total #MB before and after
* Excellent blog post by [@hd_ka](https://twitter.com/hd_ka) at [http://blog.greenbrain.de/2014/10/fire-up-those-caches.html](http://blog.greenbrain.de/2014/10/fire-up-those-caches.html)

## Quick Start
* Download `SPBestWarmup.ps1`
* Copy `SPBestWarmup.ps1` to one SharePoint web front end (WFE)
* Run `SPBestWarmup.ps1 -f` to install farm wide. Creates Task Scheduler job on every machine.
* Run `SPBestWarmup.ps1 -i` to install locally. Creates Task Scheduler job on the local machine.
* Run `SPBestWarmup.ps1 -u` to uninstall farm wide. Deletes any Task Scheduler job named "SPBestWarmup"
* Sit back and enjoy!

## Screenshots

* Install farm wide
![image](https://raw.githubusercontent.com/spjeff/spbestwarmup/master/doc/1.jpg)

* Manual run
![image](https://raw.githubusercontent.com/spjeff/spbestwarmup/master/doc/2.jpg)

* Manual run with custom URL parameter
![image](https://raw.githubusercontent.com/spjeff/spbestwarmup/master/doc/2.jpg)

## Admin Tip
* After reboot run this command to manually trigger the job and warm up IIS
`SCHTASKS /RUN /TN "SPBestWarmup"`

## Custom URLs
* Use the `-url` paramter to add custom URLs from the command line. 
* Rename Central Admin site title and edit lines `280-295` to add custom URLs within the script.  Good for lifecycle (dev, test, prod).

## Permission
* Running this with a different service account than farm might require you to first grant PowerShell access. This ensures the service account has access to run `Get-SPWebApplication` and `Get-SPServer` for detecting URLs to load. [http://technet.microsoft.com/en-us/magazine/gg490648.aspx
](http://technet.microsoft.com/en-us/magazine/gg490648.aspx)

## Contact
Please drop a line to [@spjeff](https://twitter.com/spjeff) or [spjeff@spjeff.com](mailto:spjeff@spjeff.com)
Thanks!  =)

![image](http://img.shields.io/badge/first--timers--only-friendly-blue.svg?style=flat-square)

## License

The MIT License (MIT)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.