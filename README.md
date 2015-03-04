## VirtualEngine.Compression ##
A PowerShell module for compressing and decompressing archive files using only native .NET Framework 4.5 classes.

* Create a new .zip archive.
* Add files to an existing .zip archive.
* Extract all files from a .zip archive.
* Extract individual files from a .zip archive.

Requires __Powershell 3.0 or above__ and the __.NET Framework v4.5 or higher__.

If you find it useful, unearth any bugs or have any suggestions for improvements, feel free to add an [issue](https://github.com/virtualengine/Compression/issues) or place a comment at the project home page</a>.

##### Screenshots
![ScreenShot](./VirtualEngine.Compression.png?raw=true)

##### Installation

* Automatic (via Chocolatey):
 * Run 'choco install VirtualEngine-Compression'.
 * Run 'Import-Module VirtualEngine.Compression'.
* Automatic (via OneGet on Windows 10 - until I can publish this to the PSGallery feed):
 * Run 'Install-Package VirtualEngine-Compression -Source chocolatey'.
 * Launch the PowerShell ISE.
 * Run 'Import-Module VirtualEngine.Compression'.
* Manual:
 * Download the [latest release](https://github.com/virtualengine/Compression/releases/latest).
 * Ensure the .zip file is unblocked (properties of the file / General) and extract to your Powershell module directory "$env:USERPROFILE\Documents\WindowsPowerShell\Modules".
 * Launch the PowerShell ISE.
 * Run 'Import-Module VirtualEngine.Compression'.
 * If you want it to be loaded automatically when ISE starts, add the line above to your ISE profile (see $profile).

#### Usage
Refer to the built-in cmdlet help.

* <b>Get a list of available cmdlets:</b> Get-Command -Module VirtualEngine.Compression
* <b>Get an individual's cmdlet help:</b> Get-Help New-ZipArchive -Full

##### Why?

Because we couldn't find a PowerShell module that doesn't require 3rd party assemblies and we needed the ability to build .zip archives with [PSake](https://github.com/psake/psake). In addition, this module will become a dependency for other Virtual Engine modules and DSC resources in the near future.

##### Implementation details
Written in PowerShell :)
