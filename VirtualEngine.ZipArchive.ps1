#region Public Functions

<#
.SYNOPSIS
    Extracts a Zip Archive.
.DESCRIPTION
    Extracts the entire contents of a Zip Archive.
.PARAMETER Path
    File path to the source Zip Archive (.zip) file to be extracted.
.PARAMETER LiteralPath
    Literal file path to the source Zip Archive (.zip) to be extracted.
.PARAMETER DestinationPath
    Destination directroy path to extract the Zip Archive (.zip) file contents to.
.PARAMETER Force
    By default, the Expand-ZipArchive cmdlet will not overwrite an existing file
    in the destination output directory. To overwrite existing files you must specify
    the -Force parameter.
.EXAMPLE
    Expand-ZipArchive -Path ~\Desktop\Example.zip -DestinationPath ~\Documents\Example\

    This command extracts the contents of the 'Example.zip' file on the user's desktop into
    the 'Example' directory in the user's Documents directory. Any existing files in the
    'Example' directory will not be overwritten.
.EXAMPLE
    Expand-ZipArchive -Path ~\Desktop\Example.zip -DestinationPath ~\Documents\Example\ -Force

    This command extracts the contents of the 'Example.zip' file on the user's desktop into
    the 'Example' directory in the user's Documents directory. Any existing files in the
    'Example' directory will be overwritten without warning.
.OUTPUTS
    A System.IO.FileInfo object for each extracted file.
#>
function Expand-ZipArchive {
    [CmdletBinding(DefaultParameterSetName='Path', HelpUri = 'https://github.com/VirtualEngine/Compression')]
    [OutputType([System.IO.FileInfo])]
    Param (
        # Source path to the Zip Archive.
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName='Path')]
            [ValidateNotNullOrEmpty()] [Alias('PSPath','FullName')] [string] $Path = (Get-Location -PSProvider FileSystem),
        # Source path to the Zip Archive.
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName='LiteralPath')]
            [ValidateNotNullOrEmpty()] [string[]] $LiteralPath,
        # Destination file path to extarct the Zip Archive item to.
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=1)]
            [ValidateNotNullOrEmpty()] [string] $DestinationPath,
        # Overwrite existing files
        [Switch] $Force
    )

    Begin {

        ## Validate destination path      
        if (-not(Test-Path $DestinationPath -IsValid)) { throw "Invalid Zip Archive destination path '$DestinationPath'."; }
        Write-Verbose "Resolving destination path '$DestinationPath'.";
        $DestinationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath);

        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            Write-Verbose "Resolving path(s) '$Path'.";
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);
        }
        else {
            ## Set the path to the literal path specified
            $Path = $LiteralPath;
        }

        ## If all tests passed, load the required .NET assemblies
        Write-Debug "Loading 'System.IO.Compression' .NET binaries.";
        Add-Type -AssemblyName "System.IO.Compression";
        Add-Type -AssemblyName "System.IO.Compression.FileSystem";
    } # end begin

    Process {

        foreach ($pathEntry in $Path) {

            try {

                Write-Verbose "Expanding Zip Archive '$pathEntry'.";
                $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($pathEntry);
                if ($Force) {
                    Expand-ZipArchiveItem -InputObject ([ref] $zipArchive.Entries) -DestinationPath $DestinationPath -Force;
                }
                else {
                    Expand-ZipArchiveItem -InputObject ([ref] $zipArchive.Entries) -DestinationPath $DestinationPath;
                }
     
            } # end try
            catch {
                Write-Error $_.Exception;
            }
            finally {
                ## Close the file handle
                if ($zipArchive -ne $null) { $zipArchive.Dispose(); }
            }
        } # end foreach
    } # end process

    End {
        ## Close the file handle (just in case!)
        if ($zipArchive -ne $null) { $zipArchive.Dispose(); }
    }
}

<#
.SYNOPSIS
    Extracts file(s) from a Zip Archive.
.DESCRIPTION
    The Expand-ZipArchiveItem cmdlet extracts an individual file from a Zip Archive.
.PARAMETER Path
    Internal ZipArchiveItem path inside the source Zip Archive (.zip) file to be extracted.
.PARAMETER DestinationPath
    Destination directroy path to extract the Zip Archive (.zip) file contents to.
.PARAMETER Force
    By default, the Expand-ZipArchive cmdlet will not overwrite an existing file
    in the destination output directory. To overwrite existing files you must specify
    the -Force parameter.
.EXAMPLE
    $ZipContents = Get-ZipArchiveEntry -Path ~\Desktop\Example.zip
    $ZipContents[0] | Expand-ZipArchiveItem -DestinationPath ~\Documents\Example\

    This command extracts the first item from the 'Example.zip' file located on the user's
    desktop directory. The file will be extracted to the user's Documents\Example directory.
    Any existing file in the destination direcotory will not be overwritten.
.EXAMPLE
    Expand-ZipArchiveItem -Path "Example.txt" -DestinationPath ~\Documents\Example\

    This command extracts the 'Example.txt' file from the Zip Archive to the user's \Documents\
    Example directory. Any existing file in the destination direcotory will not be overwritten.
.EXAMPLE
    Expand-ZipArchiveItem -Path "SubFolder\Example2.txt" -DestinationPath ~\Documents\Example\ -Force

    This command extracts the 'Example.txt' file from the Zip Archive 'SubFolder' directory, to
    the user's \Documents\Example directory. Any existing file in the destination directory will
    be overwritten without warning.
.INPUTS
    You can pipe System.IO.Compression.ZipArchiveEntry objects to the Expand-ZipArchiveItem
    cmdlet that are produced by the Get-ZipArchiveEntry cmdlet.
.OUTPUTS
    A System.IO.FileInfo object for each extracted file.
#>
function Expand-ZipArchiveItem {
    [CmdletBinding(DefaultParameterSetName='Path', HelpUri = 'https://github.com/VirtualEngine/Compression')]
    [OutputType([System.IO.FileInfo])]
    Param (
        # Source path to the Zip Archive.
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName='Path')]
            [ValidateNotNullOrEmpty()] [System.IO.Compression.ZipArchiveEntry[]] [ref] $InputObject,
        # Destination file path to extarct the Zip Archive item to.
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=1)]
            [ValidateNotNullOrEmpty()] [string] $DestinationPath,
        # Overwrite existing files
        [Switch] $Force  
    )

    Begin {
        ## Load the required .NET assemblies, just in case
        Write-Debug "Loading 'System.IO.Compression' .NET binaries.";
        Add-Type -AssemblyName "System.IO.Compression";
        Add-Type -AssemblyName "System.IO.Compression.FileSystem";
    }

    Process {
        
        try {

            foreach ($zipArchiveEntry in $InputObject) {

                if ($zipArchiveEntry.FullName.Contains('\')) {
                    ## We need to create the directory path as the ExtractToFile extension
                    ## method won't do this and will throw an exception

                    $pathSplit = $zipArchiveEntry.FullName.Split('\');
                    $relativeDirectoryPath = New-Object System.Text.StringBuilder;

                    ## Generate the relative directory name
                    for ($pathSplitPart = 0; $pathSplitPart -lt ($pathSplit.Count -1); $pathSplitPart++) {
                        [ref] $null = $relativeDirectoryPath.AppendFormat("{0}\", $pathSplit[$pathSplitPart]); 
                    }
         
                    ## Create the destination directory path, joining the relative directory name
                    $directoryPath = Join-Path $DestinationPath $relativeDirectoryPath.ToString().Trim('\');
                    [ref] $null = _NewDirectory -Path $directoryPath;
                        
                    $fullDestinationFilePath = Join-Path $directoryPath $zipArchiveEntry.Name;

                } # end if
                else {
                    ## Just a file in the root so just use the $DestinationPath
                    $fullDestinationFilePath = Join-Path $DestinationPath $zipArchiveEntry.Name;
                } # end else

                    ## Are we overwriting existing files (-Force)?
                if (!$Force -and (Test-Path -Path $fullDestinationFilePath -PathType Leaf)) {
                    Write-Warning "Target file '$fullDestinationFilePath' already exists.";
                }
                else {
                    ## Just overwrite any existing file
                    Write-Verbose "Extracting Zip Archive Entry '$fullDestinationFilePath'.";
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($zipArchiveEntry, $fullDestinationFilePath, $true);
                    ## Return a FileInfo object to the pipline
                    Write-Output (Get-Item -Path $fullDestinationFilePath);
                } # end if

            } # end foreach zipArchiveEntry
        } # end try
        catch {
            Write-Error $_.Exception;
        }

    } # end process
}

<#
.SYNOPSIS
    Gets content of a Zip Archive.
.DESCRIPTION
    The Get-ZipArchiveEntry cmdlet gets the file contents of a Zip Archive. The results
    of the this cmdlet can be used with the Expand-ZipArchiveItem cmdlet to extract one
    or files.
.PARAMETER Path
    File path to the source Zip Archive (.zip) file to be enumerated.
.PARAMETER LiteralPath
    Literal file path to the source Zip Archive (.zip) to be enumerated.
.EXAMPLE
    Get-ZipArchiveEntry -Path ~\Desktop\Example.zip

    This commands returns the contents of the 'Example.zip' .zip file on the user's
    desktop.
.OUTPUTS
    A System.IO.Compression.ZipArchiveEntry object per Zip Archive item.
#>
function Get-ZipArchiveEntry {
    [CmdletBinding(DefaultParameterSetName='Path', HelpUri = 'https://github.com/VirtualEngine/Compression')]
    [OutputType([System.IO.Compression.ZipArchiveEntry])]
    Param (
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName='Path')]
            [ValidateNotNullOrEmpty()] [Alias('PSPath','FullName')] [string] $Path = (Get-Location -PSProvider FileSystem),
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName='LiteralPath')]
            [ValidateNotNullOrEmpty()] [string[]] $LiteralPath
    )

    Begin {

        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            Write-Verbose "Resolving path(s) '$Path'.";
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);
            $Path = Resolve-Path $Path;
        }
        else {
            ## Set the path to the literal path specified
            $Path = $LiteralPath;
        }

        ## If all tests passed, load the required .NET assemblies
        Write-Debug "Loading 'System.IO.Compression' .NET binaries.";
        Add-Type -AssemblyName "System.IO.Compression";
        Add-Type -AssemblyName "System.IO.Compression.FileSystem";
    
    } # end begin

    Process {

        foreach ($pathEntry in $Path) {
            Write-Verbose "Processing Zip Archive '$pathEntry'.";

            try {
                $fileStream = New-Object System.IO.FileStream($pathEntry, [System.IO.FileMode]::Open);
                $zipArchive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Read);
                $zipArchive.Entries;
            }
            catch {
                Write-Error $_.Exception;
            }
            finally {
                ## Clean up
                if ($zipArchive -ne $null) { $zipArchive.Dispose(); }
                if ($fileStream -ne $null) { $fileStream.Close(); }
            }

        } # end foreach
    } # end process
}

<#
.SYNOPSIS
    Adds file(s) to an existing Zip Archive.
.DESCRIPTION
    The Add-ZipArchiveItem cmdlets adds one or more files to an existing Zip Archive.
.PARAMETER Path
    File path to the source file to be added to the Zip Archive.
.PARAMETER LiteralPath
    Absolute file path to the source file to be added to the Zip Archive.
.PARAMETER DestinationPath
    Destination Zip Archive (.zip) file to add the files to.
.PARAMETER CompressionLevel
    The compression algorithm to use. You must specify either 'Optimal', 'Fastest' or
    'NoCompression'. By default, optimal compression is used.
.PARAMETER Force
    By default, the Add-ZipArchiveItem cmdlet will not overwrite an existing file
    in the Zip Archive. To overwrite existing files within the Zip Archive, you must
    specify the -Force parameter.
.OUTPUTS
    System.IO.FileInfo
#>
function Add-ZipArchiveItem {
    [CmdletBinding(DefaultParameterSetName='Path', HelpUri = 'https://github.com/VirtualEngine/Compression')]
    [OutputType([System.IO.FileInfo])]
    Param (
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName='Path')]
            [ValidateNotNullOrEmpty()] [Alias('PSPath','FullName')] [string[]] $Path = (Get-Location -PSProvider FileSystem),
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName='LiteralPath')]
            [ValidateNotNullOrEmpty()] [string[]] $LiteralPath,
        # Existing Zip Archive file path
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=1)]
            [ValidateNotNullOrEmpty()] [string] $DestinationPath,
        # Compression level
        [Parameter(ValueFromPipelineByPropertyName=$true)]
            [ValidateSet('Optimal', 'Fastest', 'NoCompression')] [string] $CompressionLevel = 'Optimal',
        # Overwrite existing Zip Archive entries if present
        [Parameter(ValueFromPipelineByPropertyName=$true)] [Switch] $Force
    )

    Begin {

        ## Validate destination path      
        if (-not(Test-Path $DestinationPath -IsValid)) { throw "Invalid Zip Archive destination path '$DestinationPath'."; }
        Write-Verbose "Resolving destination path '$DestinationPath'.";
        $DestinationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath);
        $DestinationPath = Resolve-Path $DestinationPath;

        $resolvedPaths = @();
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            foreach ($pathItem in $Path) {
                Write-Verbose "Resolving source path(s) '$pathItem'.";
                $pathItem = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($pathItem);
                $resolvedPaths += Resolve-Path $pathItem;
            }
        }
        else {
            ## Set the path to the literal path specified
            $Path = $LiteralPath;
        }

        ## If all tests passed, load the required .NET assemblies
        Write-Debug "Loading 'System.IO.Compression' .NET binaries.";
        Add-Type -AssemblyName "System.IO.Compression";
        Add-Type -AssemblyName "System.IO.Compression.FileSystem";

        Write-Verbose "Opening existing Zip Archive '$DestinationPath'.";
        [System.IO.FileStream] $fileStream = New-Object System.IO.FileStream($DestinationPath, [System.IO.FileMode]::OpenOrCreate);
        [System.IO.Compression.ZipArchive] $zipArchive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Update);
    
    } # end begin

    Process {

        foreach ($path in $resolvedPaths) {
            if ($Force) { _ProcessZipArchivePath -Path $path -ZipArchive ([ref] $zipArchive) -Force; }
            else { _ProcessZipArchivePath -Path $path -ZipArchive ([ref] $zipArchive); }
        }
    }

    End {
        _CloseZipArchive;
    }

}

<#
.SYNOPSIS
    Creates a new Zip Archive.
.DESCRIPTION
    The New-ZipArchive cmdlet creates a new Zip Archive from files or
    directory paths passed vai the $Path parameter.
.PARAMETER Path
    File path to the source file to be added to the Zip Archive.
.PARAMETER LiteralPath
    Absolute file path to the source file to be added to the Zip Archive.
.PARAMETER DestinationPath
    Destination Zip Archive (.zip) file to add the files to.
.PARAMETER CompressionLevel
    The compression algorithm to use. You must specify either 'Optimal', 'Fastest' or
    'NoCompression'. By default, optimal compression is used.
.PARAMETER Force
    By default, the Add-ZipArchiveItem cmdlet will not overwrite an existing file
    in the Zip Archive. To overwrite existing files within the Zip Archive, you must
    specify the -Force switch parameter.
.PARAMETER NoClobber
    By default, the New-ZipArchive cmdlet will overwrite an exiting Zip Archive file,
    if present. To avoid overwriting the Zip Archive file, you must specify the
    -NoClobber switch parameter.
.EXAMPLE
    New-ZipArchive -Path .\ExampleFolder -DestinationPath ~\Desktop\Example.zip

    This command compresses the files and sub-folders of the .\ExampleFolder into a
    Zip Archive called 'Example.zip' that is placed on the user's desktop.
.OUTPUTS
    System.IO.FileInfo
#>
function New-ZipArchive {
    [CmdletBinding(DefaultParameterSetName='Path', HelpUri = 'https://github.com/VirtualEngine/Compression')]
    [OutputType([System.IO.FileInfo])]
    Param (
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName='Path')]
            [ValidateNotNullOrEmpty()] [Alias('PSPath','FullName')] [string[]] $Path = (Get-Location -PSProvider FileSystem),
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName='LiteralPath')]
            [ValidateNotNullOrEmpty()] [string[]] $LiteralPath,
        # Zip file output name
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=1)]
            [ValidateNotNullOrEmpty()] [string] $DestinationPath,
        # Compression level
        [Parameter(ValueFromPipelineByPropertyName=$true)]
            [ValidateSet('Optimal', 'Fastest', 'NoCompression')] [string] $CompressionLevel = 'Optimal',
        # Overwrite existing Zip Archive entries if present
        [Parameter(ValueFromPipelineByPropertyName=$true)] [Switch] $Force,
        # Do not create a new Zip Archive file if present
        [Parameter(ValueFromPipelineByPropertyName=$true)] [Switch] $NoClobber
    )

    Begin {

        ## Validate destination path      
        if (-not(Test-Path $DestinationPath -IsValid)) { throw "Invalid Zip Archive destination path '$DestinationPath'."; }
        Write-Verbose "Resolving destination path '$DestinationPath'.";
        $DestinationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath);

        $resolvedPaths = @();
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            foreach ($pathItem in $Path) {
                Write-Verbose "Resolving source path(s) '$pathItem'.";
                $pathItem = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($pathItem);
                $resolvedPaths += Resolve-Path $pathItem;
            }
        }
        else {
            ## Set the path to the literal path specified
            $Path = $LiteralPath;
        }      

        ## If all tests passed, load the required .NET assemblies
        Write-Debug "Loading 'System.IO.Compression' .NET binaries.";
        Add-Type -AssemblyName "System.IO.Compression";
        Add-Type -AssemblyName "System.IO.Compression.FileSystem";

        if ($NoClobber) {
            Write-Verbose "Opening an existing or creating a new Zip Archive '$DestinationPath'.";
            [System.IO.FileStream] $fileStream = New-Object System.IO.FileStream($DestinationPath, [System.IO.FileMode]::OpenOrCreate);
        }   
        else {
            ## (Re)create a new Zip Archive 
            Write-Verbose "Creating new Zip Archive '$DestinationPath'.";
            [System.IO.FileStream] $fileStream = New-Object System.IO.FileStream($DestinationPath, [System.IO.FileMode]::Create);
        }

        [System.IO.Compression.ZipArchive] $zipArchive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Update);
    
    } # end begin

    Process {

        foreach ($path in $resolvedPaths) {
            if ($Force) { _ProcessZipArchivePath -Path $path -ZipArchive ([ref] $zipArchive) -Force; }
            else { _ProcessZipArchivePath -Path $path -ZipArchive ([ref] $zipArchive); }
        }

    } # end process

    End {
        _CloseZipArchive;
        ## Return a System.IO.FileInfo to the pipeline
        Get-Item $DestinationPath;
    } # end end
}

#endregion Public Functions

#region Private Functions

<#
.Synopsis
   Creates a filesystem directory.
.DESCRIPTION
   The New-Directory cmdlet will create the target directory
   if it doesn't already exist. If the target path already exists,
   the cmdlet does nothing.
.EXAMPLE
   New-Directory -Path ~\Desktop\Example

   This example will create a folder in the user's desktop folder
   if it does not already exist.
.INPUTS
   You can pipe multiple strings or multiple System.IO.DirectoryInfo
   objects to this cmdlet.
.OUTPUTS
   System.IO.DirectoryInfo
.NOTES
    This is an internal function and should not be called directly.
#>
function _NewDirectory
{
    [CmdletBinding(DefaultParameterSetName="ByString", SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    [OutputType([System.IO.DirectoryInfo])]
    Param (
        # Target filesystem directory to create
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true,
            Position=0, ParameterSetName='ByDirectoryInfo')]
        [ValidateNotNullOrEmpty()] [System.IO.DirectoryInfo[]] $InputObject,
        
        # Target filesystem directory to create
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true,
            Position=0, ParameterSetName='ByString')] [Alias("PSPath")]
        [ValidateNotNullOrEmpty()] [string[]] $Path
    )

    Process {
        Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
        switch ($PSCmdlet.ParameterSetName)
        {
            'ByString' {
                foreach ($Directory in $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)) {
                    Write-Debug ("Testing target directory '{0}'." -f $Directory);
                    if (!(Test-Path $Directory -PathType Container)) {
                        if ($PSCmdlet.ShouldProcess($Directory, "Create directory")) {
                            Write-Verbose ("Creating target directory '{0}'." -f $Directory);
                            New-Item -Path $Directory -ItemType Directory;
                        }
                    } else {
                        Write-Debug ("Target directory '{0}' already exists." -f $Directory);
                        Get-Item -Path $Directory;
                    }
                }
            }

            'ByDirectoryInfo' {
                 foreach ($DirectoryInfo in $InputObject) {
                    Write-Debug ("Testing target directory '{0}'." -f $DirectoryInfo.FullName);
                    if (!($DirectoryInfo.Exists)) {
                        if ($PSCmdlet.ShouldProcess($DirectoryInfo.FullName, "Create directory")) {
                            Write-Verbose ("Creating target directory '{0}'." -f $DirectoryInfo.FullName);
                            New-Item -Path $DirectoryInfo.FullName -ItemType Directory;
                        }
                    } else {
                        Write-Debug ("Target directory '{0}' already exists." -f $DirectoryInfo.FullName);
                        $DirectoryInfo;
                    }
                }
            }
        }
    }
}

<#
.SYNOPSIS
    Tidies up and closes open Zip Archives and file handles
.NOTES
    This is an internal function and should not be called directly.
#>
function _CloseZipArchive {
    Process {
        ## Clean up
        Write-Verbose "Saving Zip Archive '$DestinationPath'.";
        if ($zipArchive -ne $null) { $zipArchive.Dispose(); }
        if ($fileStream -ne $null) { $fileStream.Close(); }
    }
}

<#
.SYNOPSIS
    Adds the specified paths to a Zip Archive object reference.
.NOTES
    This is an internal function and should not be called directly.
#>
function _ProcessZipArchivePath {
    Param (
        [Parameter()] [ValidateNotNullOrEmpty()] [string[]] $Path,
        [Parameter()] [ValidateNotNull()] [System.IO.Compression.ZipArchive] [ref] $ZipArchive,
        [Switch] $Force
    )

    Begin {

    }

    Process {

        foreach ($pathEntry in $Path) {
            if (Test-Path -Path $pathEntry -PathType Container) {
                ## The base directory is used for internal references to directories within the Zip Archive
                $BasePath = New-Object System.IO.DirectoryInfo($pathEntry);

                if ($Force) { [ref] $null = _AddZipArchiveItem -Path $pathEntry -ZipArchive ([ref] $zipArchive) -Force; }
                else { [ref] $null = _AddZipArchiveItem -Path $pathEntry -ZipArchive ([ref] $zipArchive); }
            }
            else {
                $fileInfo = New-Object System.IO.FileInfo($pathEntry);
                
                if (!$Force -and (_TestZipArchiveEntry -ZipArchive ([ref] $zipArchive) -Name $fileInfo.Name)) {
                    Write-Warning "Zip Archive entry '$($fileInfo.Name)' already exists.";
                }
                else {
                    Write-Verbose "Adding Zip Archive entry '$($fileInfo.Name)'.";
                    [ref] $null = _TestZipArchiveEntry -ZipArchive ([ref] $zipArchive) -Name $fileInfo.Name -Delete;
                    [ref] $null = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zipArchive, $fileInfo.FullName, $fileInfo.Name);
                }
            }
        }
    }
}

<#
.SYNOPSIS
    Tests whether a Zip Archive file contains the specified file.
.NOTES
    This is an internal function and should not be called directly.
#>
function _TestZipArchiveEntry {
    [CmdletBinding()]
    [OutputType([Boolean])]
    Param (
        # Reference to the Zip Archive object
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)] [ValidateNotNull()]
            [System.IO.Compression.ZipArchive] [ref] $ZipArchive,
        # Zip archive entry name, i.e. Subfolder\Filename.txt
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNullOrEmpty()] [string] $Name,
        # Remove zip archive entry if present
        [Switch] $Delete
    )

    Process {
        $ZipArchiveEntry = $ZipArchive.GetEntry($Name);

        if ($zipArchiveEntry -eq $null) { return $false; }
        else {
            ## Delete the entry if instructed
            if ($Delete) {
                Write-Debug "Deleting existing Zip Archive entry '$Name'.";
                $ZipArchiveEntry.Delete();
            }

            return $true;
        }
    }
}

<#
.SYNOPSIS
    Deletes a Zip Archive entry if it exists.
.NOTES
    This is an internal function and should not be called directly.
#>
function _RemoveZipArchiveEntry {
    [CmdletBinding()]
    [OutputType([Boolean])]
    Param (
        # Reference to the Zip Archive object
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNull()] [System.IO.Compression.ZipArchive] [ref] $ZipArchive,
        # Zip archive entry name, i.e. Subfolder\Filename.txt
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNullOrEmpty()] [string] $Name
    )

    Process {
        _TestZipArchiveEntry -ZipArchive ([ref] $ZipArchive) -Name $Name -Delete;
    }
}

<#
.SYNOPSIS
    Adds an item to an existing System.IO.Compression.ZipArchive.
.NOTES
    This is an internal function and should not be called directly.
#>
function _AddZipArchiveItem {
    [CmdletBinding()]
    [OutputType([System.IO.Compression.ZipArchiveEntry])]
    Param (
        # Directory path to add to the Zip Archive
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNullOrEmpty()] [string] $Path,
        # Reference to the ZipArchive object
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNull()] [System.IO.Compression.ZipArchive] [ref] $ZipArchive,
        # Base directory path
        [Parameter(ValueFromPipelineByPropertyName=$true)]
            [AllowNull()] [string] $BasePath = '',
        # Overwrite existing Zip Archive entries if present
        [Switch] $Force
    )

    Process {
        Write-Debug "Resolving directory path '$Path'.";
        foreach ($childItem in (Get-ChildItem -Path $Path)) {
            
            if (Test-Path -Path $childItem.FullName -PathType Container) {
                ## Recurse subfolder, expanding the base directory, i.e. SubFolder1\SubFolder2
                if ([string]::IsNullOrEmpty($BasePath)) { $newBasePath = New-Object System.IO.DirectoryInfo($childItem).Name; }
                else { $newBasePath = "$BasePath\$((New-Object System.IO.DirectoryInfo($childItem)).Name)"; }

                if ($Force) { _AddZipArchiveItem -Path $childItem.FullName -ZipArchive ([ref]$ZipArchive) -BasePath $newBasePath -Force; }
                else { _AddZipArchiveItem -Path $childItem.FullName -ZipArchive ([ref]$ZipArchive) -BasePath $newBasePath; }
            } # end if
            else {
                ## Add the file using the current base directory
                if ([string]::IsNullOrEmpty($BasePath)) { $childItemPath = $childItem; }
                else { $childItemPath = "$BasePath\$childItem"; }

                if (!$Force -and (_TestZipArchiveEntry -ZipArchive ([ref] $zipArchive) -Name $childItemPath)) {
                    Write-Warning "Zip Archive entry '$childItemPath' already exists.";
                }
                else {
                    Write-Verbose "Adding Zip Archive entry '$childItemPath'.";
                    [ref] $null = _TestZipArchiveEntry -ZipArchive ([ref] $zipArchive) -Name $childItemPath -Delete;
                    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zipArchive, $childItem.FullName, $childItemPath);
                }
            } # end else

        } # end foreach
    } # end process
}

#endregion Private Functions