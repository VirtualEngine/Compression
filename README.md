## VirtualEngine.Compression ##
A PowerShell module for compressing and decompressing archive files using only native .NET Framework 4.5 classes.

* Create a .zip archive.
* Add files to an existing .zip archive.
* Extract all files from a .zip archive.
* Extract individual files from a .zip archive.

Requires __Powershell 3.0 or above__ and the __.NET Framework v4.5 or higher__.

If you find it useful, unearth any bugs or have any suggestions for improvements, feel free to add an <a href="https://github.com/virtualengine/Compression/issues">issue</a> or place a comment at the project home page</a>.

##### Screenshots
![ScreenShot](./VirtualEngine.Compression.png?raw=true)

##### Installation

* Automatic (via Chocolatey):
 * Run 'cinst VirtualEngine-Compression'.
 * Run 'Import-Module VirtualEngine-Compression'.
* Manual:
 * Download the latest release .zip.
 * Extract the .zip to your somewhere in the $PSModulePath, e.g. \Document\WindowsPowerShell\Modules\.
 * Run 'Import-Module VirtualEngine-Compression'.
 * If you want it to be loaded automatically when PowerShell starts, add the line above to your PowerShell profile (see $profile).

#### Usage
Refer to the built-in cmdlet help.

* <b>Get a list of available cmdlets:</b> Get-Command -Module VirtualEngine.Compression
* <b>Get an individual's cmdlet help:</b> Get-Help New-ZipArchive -Full

##### Why?

Because we couldn't find a PowerShell module that doesn't require 3rd party assemblies and we needed the ability to build .zip archives with <a href="https://github.com/psake/psake">Psake</a>. In addition, this module will become a dependency for other Virtual Engine modules and DSC resources in the near future.

##### Implementation details
Written in PowerShell :)
