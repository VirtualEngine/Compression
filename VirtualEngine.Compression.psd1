@{
    RootModule = 'VirtualEngine.Compression.psm1';
    ModuleVersion = '1.1.0';
    GUID = 'fa270b73-f196-4e26-9ce6-4ae04d7fbe79';
    Author = 'Iain Brighton';
    CompanyName = 'Virtual Engine';
    Copyright = '(c) 2015 Virtual Engine Limited. All rights reserved.';
    Description = 'Virtual Engine Zip archive compression PowerShell cmdlets.';
    PowerShellVersion = '3.0';
    DotNetFrameworkVersion = '4.5';
    CLRVersion = '4.0';
    FunctionsToExport = '*-*';
    FileList = @('VirtualEngine.Compression.psm1','VirtualEngine.ZipArchive.ps1','VirtualEngine.Compression.psd1','LICENSE');
    PrivateData = @{
        PSData = @{
            Tags = @('VirtualEngine','Powershell','Compression','ZIP','Archive');
            LicenseUri = 'https://raw.githubusercontent.com/VirtualEngine/Compression/master/LICENSE';
            ProjectUri = 'https://github.com/VirtualEngine/Compression';
            IconUri = 'https://cdn.rawgit.com/VirtualEngine/Compression/38aa3a3c879fd6564d659d41bffe62ec91fb47ab/icon.png';
        } # End of PSData hashtable
    } # End of PrivateData hashtable# Private data to pass to the module specified in RootModule/ModuleToProcess

    HelpInfoURI = 'https://github.com/VirtualEngine/Compression';
}
