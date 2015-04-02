## Import the .Net 4.5 compression binaries as this is needed to be able to parse the
## VirtualEngine.ZipArchive.ps1 file.
Write-Debug 'Loading ''System.IO.Compression'' .NET binaries.';
Add-Type -AssemblyName 'System.IO.Compression';
Add-Type -AssemblyName 'System.IO.Compression.FileSystem';

## Import the VirtualEngine.ZipArchive.ps1 file. This permits loading of the module's
## functions for unit testing, without having to unload/load the whole module.
. (Join-Path -Path (Split-Path -Path $PSCommandPath) -ChildPath VirtualEngine.ZipArchive.ps1);

## Export public functions
Export-ModuleMember -Function *-*;
