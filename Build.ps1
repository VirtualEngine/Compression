$Properties = @{
    CertificateThumbprint = 'D10BB31E5CE3048A7D4DA0A4DD681F05A85504D3';
    ReleaseDirectory = '.\Releases';
    TimeStampServer = 'http://timestamp.verisign.com/scripts/timestamp.dll';
    LicenseUrl = 'https://raw.githubusercontent.com/VirtualEngine/ZipArchive/master/LICENSE';
    DownloadBaseUrl = 'http://virtualengine.co.uk/wp-content/uploads';
}

Import-Module VirtualEngine.Build -ErrorAction Stop;
#Import-Module Posh-SSH -ErrorAction Stop;

$module = Get-ModuleManifest;
$packageName = "$($module.Name.Replace('.','-'))-$($module.Version.ToString())";
$tempDirectory = Join-Path $env:TEMP $module.Name;
$zipDirectory = Join-Path $tempDirectory $module.Name;

task default -depends CreateReleaseZip, CreateChocolateyReleasePackage, RemoveReleaseDirectory

task CreateReleaseZipDirectory {
    Remove-Item -Path $tempDirectory -Recurse -Force -ErrorAction SilentlyContinue;
    [ref] $null = New-Item -Path $zipDirectory -ItemType Container;
}

task StageReleaseZipFiles -depends CreateReleaseZipDirectory {

    $codeSigningCert = Get-ChildItem Cert:\ -CodeSigningCert -Recurse | Where Thumbprint -eq $Properties.CertificateThumbprint;

    foreach ($moduleFile in Get-ModuleFiles) {
        Copy-Item -Path $moduleFile.FullName -Destination $zipDirectory -Force;

        if ($moduleFile.Extension -in '.ps1','.psm1') {
            $moduleFilePath = Join-Path $zipDirectory $moduleFile.Name;
            Write-Verbose ("Signing file '{0}'." -f $moduleFilePath);
            $signResult = Set-ScriptSigntaure -Path $moduleFilePath -Thumbprint $Properties.CertificateThumbprint -TimeStampServer $Properties.TimeStampServer;
        }
    }
}

task CreateReleaseZip -depends StageReleaseZipFiles {
    $releaseDirectory = Resolve-Path (Join-Path . $Properties.ReleaseDirectory);
    $zipFileName = Join-Path $releaseDirectory ("{0}.zip" -f $packageName);
    Write-Verbose ("Zip release path '{0}'." -f $zipFileName);
    $zipFile = New-ZipArchive -Path $tempDirectory -DestinationPath $zipFileName;
    Write-Verbose ("Zip archive '{0}' created." -f $zipFile.FullName);
}

task CreateChocolateyReleaseDirectory {
    Remove-Item -Path $tempDirectory -Recurse -Force -ErrorAction SilentlyContinue;
    [ref] $null = New-Item -Path $tempDirectory -ItemType Container;
    [ref] $null = New-Item -Path "$tempDirectory\tools" -ItemType Container;
}

task StageChocolateyReleaseFiles -depends CreateChocolateyReleaseDirectory {
    ## Create .nuspec
    $nuspec = $module | New-NugetNuspec -LicenseUrl $Properties.LicenseUrl;
    $nuspecFilename = "$($module.Name).nuspec";
    $nuspecPath = Join-Path $tempDirectory $nuspecFilename;
    $nuspec.Save($nuspecPath);

    ## Create \Tools\ChocolateyInstall.ps1
    $chocolateyInstallPath = Join-Path (Get-Location) 'ChocolateyInstall.ps1'; 
    Copy-Item -Path $chocolateyInstallPath -Destination "$tempDirectory\tools\" -Force;

    ## Add Install-ChocolateyZipPackage to the ChocolateyInstall.ps1 file with the relevant download link
    $downloadUrl = "$($Properties.DownloadBaseUrl)/$packageName.zip";
    $installChocolateyZipPackage = "Install-ChocolateyZipPackage '{0}' '{1}' '$userPSModulePath';" -f $packageName, $downloadUrl;
    Add-Content -Path "$tempDirectory\tools\ChocolateyInstall.ps1" -Value $installChocolateyZipPackage;
}

task CreateChocolateyReleasePackage -depends StageChocolateyReleaseFiles {
    
    $releaseDirectory = Resolve-Path (Join-Path . $Properties.ReleaseDirectory);
    Push-Location $tempDirectory;
    Invoke-Expression -Command ('Nuget Pack "{0}" -OutputDirectory "{1}"' -f $nuspecFileName, $releaseDirectory);
    Pop-Location;
}

task PushReleaseZip -depends CreateReleaseZip {
    Import-Module Posh-SSH;

}

task RemoveReleaseDirectory {
    Remove-Item -Path $tempDirectory -Recurse -Force -ErrorAction SilentlyContinue;
}