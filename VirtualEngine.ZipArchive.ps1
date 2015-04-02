#region Public Functions

function Expand-ZipArchive {
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
    [CmdletBinding(DefaultParameterSetName='Path', HelpUri = 'https://github.com/VirtualEngine/Compression')]
    [OutputType([System.IO.FileInfo])]
    param (
        # Source path to the Zip Archive.
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position =0 , ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()] [Alias('PSPath','FullName')] [System.String] $Path = (Get-Location -PSProvider FileSystem),
        # Source path to the Zip Archive.
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 0, ParameterSetName = 'LiteralPath')]
        [ValidateNotNullOrEmpty()] [System.String[]] $LiteralPath,
        # Destination file path to extarct the Zip Archive item to.
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 1)]
        [ValidateNotNullOrEmpty()] [System.String] $DestinationPath,
        # Overwrite existing files
        [Switch] $Force
    )
    begin {
        ## Validate destination path      
        if (-not(Test-Path -Path $DestinationPath -IsValid)) {
            throw ('Invalid Zip Archive destination path ''{0}''.' -f $DestinationPath);
        }
        Write-Verbose ('Resolving destination path ''{0}''.' -f $DestinationPath);
        $DestinationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath);
        if (-not (Test-Path -Path $DestinationPath -PathType Container)) {
            Write-Verbose ('Creating destination path directory ''{0}''.' -f $DestinationPath);
            [Ref] $null = New-Item -Path $DestinationPath -ItemType Directory;
        }
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            Write-Verbose ('Resolving source path(s) ''{0}''.' -f $Path);
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);
        }
        else {
            ## Set the path to the literal path specified
            $Path = $LiteralPath;
        }
        ## If all tests passed, load the required .NET assemblies
        Write-Debug 'Loading ''System.IO.Compression'' .NET binaries.';
        Add-Type -AssemblyName 'System.IO.Compression';
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem';
    } # end begin
    process {
        foreach ($pathEntry in $Path) {
            try {
                Write-Verbose ('Expanding Zip Archive ''{0}''.' -f $pathEntry);
                $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($pathEntry);
                Expand-ZipArchiveItem -InputObject ([Ref] $zipArchive.Entries) -DestinationPath $DestinationPath -Force:$Force;
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
    end {
        ## Close the file handle (just in case!)
        if ($zipArchive -ne $null) { $zipArchive.Dispose(); }
    }
} # end function Expand-ZipArchive

function Expand-ZipArchiveItem {
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
        $ZipContents = Get-ZipArchiveItem -Path ~\Desktop\Example.zip
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
        cmdlet that are produced by the Get-ZipArchiveItem cmdlet.
    .OUTPUTS
        A System.IO.FileInfo object for each extracted file.
#>
    [CmdletBinding(DefaultParameterSetName='Path', HelpUri = 'https://github.com/VirtualEngine/Compression')]
    [OutputType([System.IO.FileInfo])]
    param (
        # Source path to the Zip Archive.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()] [System.IO.Compression.ZipArchiveEntry[]] [Ref] $InputObject,
        # Destination file path to extarct the Zip Archive item to.
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 1)]
        [ValidateNotNullOrEmpty()] [System.String] $DestinationPath,
        # Overwrite existing files
        [Switch] $Force  
    )
    begin {
        ## Load the required .NET assemblies, just in case
        Write-Debug 'Loading ''System.IO.Compression'' .NET binaries.';
        Add-Type -AssemblyName 'System.IO.Compression';
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem';
    }
    process {
        try {
            foreach ($zipArchiveEntry in $InputObject) {
                if ($zipArchiveEntry.FullName.Contains('/')) {
                    ## We need to create the directory path as the ExtractToFile extension
                    ## method won't do this and will throw an exception
                    $pathSplit = $zipArchiveEntry.FullName.Split('/');
                    $relativeDirectoryPath = New-Object System.Text.StringBuilder;

                    ## Generate the relative directory name
                    for ($pathSplitPart = 0; $pathSplitPart -lt ($pathSplit.Count -1); $pathSplitPart++) {
                        [Ref] $null = $relativeDirectoryPath.AppendFormat('{0}\', $pathSplit[$pathSplitPart]); 
                    }
         
                    ## Create the destination directory path, joining the relative directory name
                    $directoryPath = Join-Path $DestinationPath $relativeDirectoryPath.ToString().Trim('\');
                    [Ref] $null = _NewDirectory -Path $directoryPath;
                        
                    $fullDestinationFilePath = Join-Path $directoryPath $zipArchiveEntry.Name;
                } # end if
                else {
                    ## Just a file in the root so just use the $DestinationPath
                    $fullDestinationFilePath = Join-Path $DestinationPath $zipArchiveEntry.Name;
                } # end else

                if ([string]::IsNullOrEmpty($zipArchiveEntry.Name)) {
                    ## This is a folder and we need to create the directory path as the
                    ## ExtractToFile extension method won't do this and will throw an exception
                    $pathSplit = $zipArchiveEntry.FullName.Split('/');
                    $relativeDirectoryPath = New-Object System.Text.StringBuilder;

                    ## Generate the relative directory name
                    for ($pathSplitPart = 0; $pathSplitPart -lt ($pathSplit.Count -1); $pathSplitPart++) {
                        [Ref] $null = $relativeDirectoryPath.AppendFormat('{0}\', $pathSplit[$pathSplitPart]); 
                    }
         
                    ## Create the destination directory path, joining the relative directory name
                    $directoryPath = Join-Path $DestinationPath $relativeDirectoryPath.ToString().Trim('\');
                    [Ref] $null = _NewDirectory -Path $directoryPath;
                        
                    $fullDestinationFilePath = Join-Path $directoryPath $zipArchiveEntry.Name;
                }
                elseif (!$Force -and (Test-Path -Path $fullDestinationFilePath -PathType Leaf)) {
                    ## Are we overwriting existing files (-Force)?
                    Write-Warning ('Target file ''{0}'' already exists.' -f $fullDestinationFilePath);
                }
                else {
                    ## Just overwrite any existing file
                    Write-Verbose ('Extracting Zip Archive Entry ''{0}''.' -f $fullDestinationFilePath);
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
} # end function Expand-ZipArchiveItem

function Get-ZipArchiveItem {
<#
    .SYNOPSIS
        Gets the contents of a Zip Archive.
    .DESCRIPTION
        The Get-ZipArchiveItem cmdlet gets the file contents of a Zip Archive. The results
        of the this cmdlet can be used with the Expand-ZipArchiveItem cmdlet to extract one
        or files.
    .PARAMETER Path
        File path to the source Zip Archive (.zip) file to be enumerated.
    .PARAMETER LiteralPath
        Literal file path to the source Zip Archive (.zip) to be enumerated.
    .EXAMPLE
        Get-ZipArchiveItem -Path ~\Desktop\Example.zip

        This commands returns the contents of the 'Example.zip' .zip file on the user's
        desktop.
    .OUTPUTS
        A System.IO.Compression.ZipArchiveEntry object per Zip Archive item.
#>
    [CmdletBinding(DefaultParameterSetName = 'Path', HelpUri = 'https://github.com/VirtualEngine/Compression')]
    [OutputType([System.IO.Compression.ZipArchiveEntry])]
    param (
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()] [Alias('PSPath','FullName')] [System.String] $Path = (Get-Location -PSProvider FileSystem),
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 0, ParameterSetName = 'LiteralPath')]
        [ValidateNotNullOrEmpty()] [System.String[]] $LiteralPath
    )
    begin {
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            Write-Verbose ('Resolving path(s) ''{0}''.' -f $Path);
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);
            $Path = Resolve-Path -Path $Path;
        }
        else {
            ## Set the path to the literal path specified
            $Path = $LiteralPath;
        }
        ## If all tests passed, load the required .NET assemblies
        Write-Debug 'Loading ''System.IO.Compression'' .NET binaries.';
        Add-Type -AssemblyName 'System.IO.Compression';
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem';
    } # end begin
    process {
        foreach ($pathEntry in $Path) {
            Write-Verbose ('Processing Zip Archive ''{0}''.' -f $pathEntry);
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
} # end function Get-ZipArchiveItem

function Add-ZipArchiveItem {
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
    [CmdletBinding(DefaultParameterSetName = 'Path', HelpUri = 'https://github.com/VirtualEngine/Compression')]
    [OutputType([System.IO.FileInfo])]
    param (
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()] [Alias('PSPath','FullName')] [System.String[]] $Path = (Get-Location -PSProvider FileSystem),
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 0, ParameterSetName = 'LiteralPath')]
        [ValidateNotNullOrEmpty()] [System.String[]] $LiteralPath,
        # Existing Zip Archive file path
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 1)]
        [ValidateNotNullOrEmpty()] [System.String] $DestinationPath,
        # Compression level
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateSet('Optimal', 'Fastest', 'NoCompression')] [System.String] $CompressionLevel = 'Optimal',
        # Overwrite existing Zip Archive entries if present
        [Parameter(ValueFromPipelineByPropertyName=$true)] [Switch] $Force
    )
    begin {
        ## Validate destination path      
        if (-not (Test-Path -Path $DestinationPath -IsValid)) {
            throw ('Invalid Zip Archive destination path ''{0}''.' -f $DestinationPath);
        }
        Write-Verbose ('Resolving destination path ''{0}''.' -f $DestinationPath);
        $DestinationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath);
        $DestinationPath = Resolve-Path -Path $DestinationPath;
        $resolvedPaths = @();
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            foreach ($pathItem in $Path) {
                Write-Verbose ('Resolving source path(s) ''{0}''.' -f $pathItem);
                $pathItem = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($pathItem);
                $resolvedPaths += Resolve-Path -Path $pathItem;
            }
        }
        else {
            ## Set the path to the literal path specified
            $Path = $LiteralPath;
        }
        ## If all tests passed, load the required .NET assemblies
        Write-Debug 'Loading ''System.IO.Compression'' .NET binaries.';
        Add-Type -AssemblyName 'System.IO.Compression';
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem';
        Write-Verbose ('Opening existing Zip Archive ''{0}''.' -f $DestinationPath);
        [System.IO.FileStream] $fileStream = New-Object System.IO.FileStream($DestinationPath, [System.IO.FileMode]::OpenOrCreate);
        [System.IO.Compression.ZipArchive] $zipArchive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Update);
    } # end begin
    process {
        foreach ($path in $resolvedPaths) {
            _ProcessZipArchivePath -Path $path -ZipArchive ([Ref] $zipArchive) -Force:$Force;
        } # end foreach
    } # end process
    end {
        _CloseZipArchive;
    }
} #end function Add-ZipArchiveItem

function New-ZipArchive {
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
    [CmdletBinding(DefaultParameterSetName='Path', HelpUri = 'https://github.com/VirtualEngine/Compression')]
    [OutputType([System.IO.FileInfo])]
    param (
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()] [Alias('PSPath','FullName')] [System.String[]] $Path = (Get-Location -PSProvider FileSystem),
        # Source path/files to add to the .ZIP file
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 0, ParameterSetName = 'LiteralPath')]
        [ValidateNotNullOrEmpty()] [System.String[]] $LiteralPath,
        # Zip file output name
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 1)]
        [ValidateNotNullOrEmpty()] [System.String] $DestinationPath,
        # Compression level
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateSet('Optimal', 'Fastest', 'NoCompression')] [System.String] $CompressionLevel = 'Optimal',
        # Overwrite existing Zip Archive entries if present
        [Parameter(ValueFromPipelineByPropertyName = $true)] [Switch] $Force,
        # Do not create a new Zip Archive file if present
        [Parameter(ValueFromPipelineByPropertyName = $true)] [Switch] $NoClobber
    )
    begin {
        ## Validate destination path      
        if (-not (Test-Path -Path $DestinationPath -IsValid)) {
            throw ('Invalid Zip Archive destination path ''{0}''.' -f $DestinationPath);
        }
        Write-Verbose ('Resolving destination path ''{0}''.' -f $DestinationPath);
        $DestinationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath);
        $resolvedPaths = @();
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            foreach ($pathItem in $Path) {
                Write-Verbose ('Resolving source path(s) ''{0}''.' -f $pathItem);
                $pathItem = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($pathItem);
                $resolvedPaths += Resolve-Path -Path $pathItem;
            }
        }
        else {
            ## Set the path to the literal path specified
            $Path = $LiteralPath;
        }      
        ## If all tests passed, load the required .NET assemblies
        Write-Debug 'Loading ''System.IO.Compression'' .NET binaries.';
        Add-Type -AssemblyName 'System.IO.Compression';
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem';
        if ($NoClobber) {
            Write-Verbose ('Opening an existing or creating a new Zip Archive ''{0}''.' -f $DestinationPath);
            [System.IO.FileStream] $fileStream = New-Object System.IO.FileStream($DestinationPath, [System.IO.FileMode]::OpenOrCreate);
        }   
        else {
            ## (Re)create a new Zip Archive 
            Write-Verbose ('Creating new Zip Archive ''{0}''.' -f $DestinationPath);
            [System.IO.FileStream] $fileStream = New-Object System.IO.FileStream($DestinationPath, [System.IO.FileMode]::Create);
        }
        [System.IO.Compression.ZipArchive] $zipArchive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Update);
    } # end begin
    process {
        foreach ($path in $resolvedPaths) {
            _ProcessZipArchivePath -Path $path -ZipArchive ([Ref] $zipArchive) -Force:$Force;
        }
    } # end process
    end {
        _CloseZipArchive;
        ## Return a System.IO.FileInfo to the pipeline
        Get-Item -Path $DestinationPath;
    } # end end
}

#endregion Public Functions

#region Private Functions

function _NewDirectory {
<#
    .SYNOPSIS
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
    [CmdletBinding(DefaultParameterSetName = 'ByString', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    [OutputType([System.IO.DirectoryInfo])]
    param (
        # Target filesystem directory to create
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true,
            Position = 0, ParameterSetName = 'ByDirectoryInfo')]
        [ValidateNotNullOrEmpty()] [System.IO.DirectoryInfo[]] $InputObject,
        # Target filesystem directory to create
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true,
            Position = 0, ParameterSetName = 'ByString')] [Alias('PSPath')]
        [ValidateNotNullOrEmpty()] [System.String[]] $Path
    )
    process {
        Write-Debug ('Using parameter set ''{0}''.' -f $PSCmdlet.ParameterSetName);
        switch ($PSCmdlet.ParameterSetName) {
            'ByString' {
                foreach ($Directory in $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)) {
                    Write-Debug ('Testing target directory ''{0}''.' -f $Directory);
                    if (-not (Test-Path -Path $Directory -PathType Container)) {
                        if ($PSCmdlet.ShouldProcess($Directory, 'Create directory')) {
                            Write-Verbose ('Creating target directory ''{0}''.' -f $Directory);
                            New-Item -Path $Directory -ItemType Directory;
                        }
                    }
                    else {
                        Write-Debug ('Target directory ''{0}'' already exists.' -f $Directory);
                        Get-Item -Path $Directory;
                    }
                } # end foreach
            } # end ByString

            'ByDirectoryInfo' {
                 foreach ($DirectoryInfo in $InputObject) {
                    Write-Debug ('Testing target directory ''{0}''.' -f $DirectoryInfo.FullName);
                    if (!($DirectoryInfo.Exists)) {
                        if ($PSCmdlet.ShouldProcess($DirectoryInfo.FullName, 'Create directory')) {
                            Write-Verbose ('Creating target directory ''{0}''.' -f $DirectoryInfo.FullName);
                            New-Item -Path $DirectoryInfo.FullName -ItemType Directory;
                        }
                    }
                    else {
                        Write-Debug ('Target directory ''{0}'' already exists.' -f $DirectoryInfo.FullName);
                        $DirectoryInfo;
                    }
                } # end foreach
            } #end ByDirectoryInfo
        }
    } # end process
} # end function _NewDirectory

function _CloseZipArchive {
<#
    .SYNOPSIS
        Tidies up and closes open Zip Archives and file handles
    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    param ()

    process {
        ## Clean up
        Write-Verbose ('Saving Zip Archive ''{0}''.' -f $DestinationPath);
        if ($zipArchive -ne $null) {
            $zipArchive.Dispose();
        }
        if ($fileStream -ne $null) {
            $fileStream.Close();
        }
    } # end process
} # end function _CloseZipArchive

function _ProcessZipArchivePath {
<#
    .SYNOPSIS
        Adds the specified paths to a Zip Archive object Reference.
    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    param (
        [Parameter()] [ValidateNotNullOrEmpty()] [System.String[]] $Path,
        [Parameter()] [ValidateNotNull()] [System.IO.Compression.ZipArchive] [Ref] $ZipArchive,
        [Switch] $Force
    )
    process {
        foreach ($pathEntry in $Path) {
            if (Test-Path -Path $pathEntry -PathType Container) {
                ## The base directory is used for internal References to directories within the Zip Archive
                $BasePath = New-Object System.IO.DirectoryInfo($pathEntry);
                [Ref] $null = _AddZipArchiveItem -Path $pathEntry -ZipArchive ([Ref] $zipArchive) -Force:$Force;
            } # end if
            else {
                $fileInfo = New-Object System.IO.FileInfo($pathEntry);
                
                if ((-not $Force) -and (_TestZipArchiveEntry -ZipArchive ([Ref] $zipArchive) -Name $fileInfo.Name)) {
                    Write-Warning ('Zip Archive entry ''{0}'' already exists.' -f $fileInfo.Name);
                }
                else {
                    Write-Verbose ('Adding Zip Archive entry ''{0}''.' -f $fileInfo.Name);
                    [Ref] $null = _TestZipArchiveEntry -ZipArchive ([Ref] $zipArchive) -Name $fileInfo.Name -Delete;
                    [Ref] $null = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zipArchive, $fileInfo.FullName, $fileInfo.Name);
                }
            } # end else
        } # end foreach
    } # end process
} # end function _ProcessZipArchivePath

function _TestZipArchiveEntry {
<#
    .SYNOPSIS
        Tests whether a Zip Archive file contains the specified file.
    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (
        # Reference to the Zip Archive object
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNull()] [System.IO.Compression.ZipArchive] [Ref] $ZipArchive,
        # Zip archive entry name, i.e. Subfolder\Filename.txt
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()] [System.String] $Name,
        # Remove zip archive entry if present
        [Switch] $Delete
    )
    process {
        $ZipArchiveEntry = $ZipArchive.GetEntry($Name);
        if ($zipArchiveEntry -eq $null) {
            return $false;
        }
        else {
            ## Delete the entry if instructed
            if ($Delete) {
                Write-Debug ('Deleting existing Zip Archive entry ''{0}''.' -f $Name);
                $ZipArchiveEntry.Delete();
            }
            return $true;
        } # end else
    } # end process
} # end function _TestZipArchiveEntry

function _RemoveZipArchiveEntry {
<#
    .SYNOPSIS
        Deletes a Zip Archive entry if it exists.
    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (
        # Reference to the Zip Archive object
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNull()] [System.IO.Compression.ZipArchive] [Ref] $ZipArchive,
        # Zip archive entry name, i.e. Subfolder\Filename.txt
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()] [System.String] $Name
    )
    process {
        _TestZipArchiveEntry -ZipArchive ([Ref] $ZipArchive) -Name $Name -Delete;
    }
} # end _RemoveZipArchiveEntry

function _AddZipArchiveItem {
<#
    .SYNOPSIS
        Adds an item to an existing System.IO.Compression.ZipArchive.
    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [OutputType([System.IO.Compression.ZipArchiveEntry])]
    param (
        # Directory path to add to the Zip Archive
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()] [System.String] $Path,
        # Reference to the ZipArchive object
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNull()] [System.IO.Compression.ZipArchive] [Ref] $ZipArchive,
        # Base directory path
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [AllowNull()] [System.String] $BasePath = '',
        # Overwrite existing Zip Archive entries if present
        [Switch] $Force
    )
    process {
        Write-Debug ('Resolving directory path ''{0}''.' -f $Path);
        foreach ($childItem in (Get-ChildItem -Path $Path)) {
            if (Test-Path -Path $childItem.FullName -PathType Container) {
                ## Recurse subfolder, expanding the base directory, i.e. SubFolder1\SubFolder2
                if ([string]::IsNullOrEmpty($BasePath)) {
                    $newBasePath = New-Object System.IO.DirectoryInfo($childItem).Name;
                }
                else {
                    $newBasePath = '{0}\{1}' -f $BasePath, (New-Object System.IO.DirectoryInfo($childItem)).Name;
                }
            } # end if
            else {
                ## Add the file using the current base directory
                if ([string]::IsNullOrEmpty($BasePath)) {
                    $childItemPath = $childItem;
                }
                else {
                    $childItemPath = '{0}\{1}' -f $BasePath, $childItem;
                }
                if (!$Force -and (_TestZipArchiveEntry -ZipArchive ([Ref] $zipArchive) -Name $childItemPath)) {
                    Write-Warning ('Zip Archive entry ''{0}'' already exists.' -f $childItemPath);
                }
                else {
                    Write-Verbose ('Adding Zip Archive entry ''{0}''.' -f $childItemPath);
                    [Ref] $null = _TestZipArchiveEntry -ZipArchive ([Ref] $zipArchive) -Name $childItemPath -Delete;
                    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zipArchive, $childItem.FullName, $childItemPath);
                }
            } # end else

        } # end foreach
    } # end process
} # end function _AddZipArchiveItem

#endregion Private Functions
