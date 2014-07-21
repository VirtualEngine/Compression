## Load the required .NET assemblies
Add-Type -AssemblyName "System.IO.Compression";
Add-Type -AssemblyName "System.IO.Compression.FileSystem";

## Ensure the latest module version is loaded. NOTE: This replaces the
## default dot sourcing used by Pester.
Remove-Module VirtualEngine.Compression -ErrorAction SilentlyContinue;
Import-Module "..\VirtualEngine.Compression" -ErrorAction Stop;

#$here = Split-Path -Parent $MyInvocation.MyCommand.Path
#$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".")
#. "$here\$sut"

Describe "New-ZipArchive" {

    ## Setup test drive
    $testDirectory = (New-TestDrive -PassThru).Root;

    It "Creates the test directory" {

        Setup -Dir "NewZipArchive" -PassThru | Should Exist;
        $File1 = Setup -File "NewZipArchive\File1.txt" "This is test file 1." -PassThru;
        $File2 = Setup -File "NewZipArchive\File2.txt" "This is test file 2." -PassThru;
        $File3 = Setup -File "NewZipArchive\File3.txt" "This is test file 3." -PassThru;
        
        (Get-Item -Path "TestDrive:\NewZipArchive").GetType() | Should Be System.IO.DirectoryInfo;
        (Get-ChildItem -Path "TestDrive:\NewZipArchive").Count | Should Be 3;  
	}


    It "Creates 'NewZipArchive.zip' from directory" {
        $NewZipArchive = New-ZipArchive -Path "TestDrive:\NewZipArchive" -DestinationPath "$testDirectory\NewZipArchive.zip";
        (Get-Item -Path "TestDrive:\NewZipArchive.zip" -ErrorAction SilentlyContinue) | Should Not Be $null;
    }
    
    It "Checks 'NewZipArchive.zip' size" {
        (Get-Item -Path "TestDrive:\NewZipArchive.zip").Length | Should Be 364;
    }

    It "Creates 'NewZipArchive.zip' from string array" {
        $files = Get-ChildItem -Path "TestDrive:\NewZipArchive\*" -Recurse;
        $NewZipArchive = New-ZipArchive -Path $files.FullName -DestinationPath "$testDirectory\NewZipArchive.zip";
        (Get-Item -Path "TestDrive:\NewZipArchive.zip" -ErrorAction SilentlyContinue) | Should Not Be $null;
    }

    It "Checks 'NewZipArchive.zip' size" {
        (Get-Item -Path "TestDrive:\NewZipArchive.zip").Length | Should Be 364;
    }
}

Describe "Get-ZipArchiveEntry" {

    ## Setup test drive
    $testDirectory = (New-TestDrive -PassThru).Root;

    It "Creates the test directory" {

        Setup -Dir "NewZipArchive" -PassThru | Should Exist;
        $File1 = Setup -File "NewZipArchive\File1.txt" "This is test file 1." -PassThru;
        $File2 = Setup -File "NewZipArchive\File2.txt" "This is test file 2." -PassThru;
        $File3 = Setup -File "NewZipArchive\File3.txt" "This is test file 3." -PassThru;
        
        (Get-Item -Path "TestDrive:\NewZipArchive").GetType() | Should Be System.IO.DirectoryInfo;
        (Get-ChildItem -Path "TestDrive:\NewZipArchive").Count | Should Be 3;
	}


    It "Creates 'NewZipArchive.zip'" {
        $NewZipArchive = New-ZipArchive -Path "TestDrive:\NewZipArchive" -DestinationPath "$testDirectory\NewZipArchive.zip";
        (Get-Item -Path "TestDrive:\NewZipArchive.zip" -ErrorAction SilentlyContinue) | Should Not Be $null;
    }
    
    It "Ensures the Zip Archive contents" {
        $zipArchiveItems = Get-ZipArchiveEntry -Path "TestDrive:\NewZipArchive.zip";
        $zipArchiveItems.Count | Should Be 3;
    }

}

Describe "Expand-ZipArchive" {

    ## Setup test drive
    $testDirectory = (New-TestDrive -PassThru).Root;

    It "Creates the test directory" {
        Setup -Dir "NewZipArchive" -PassThru | Should Exist;
        $File1 = Setup -File "NewZipArchive\File1.txt" "This is test file 1." -PassThru;
        $File2 = Setup -File "NewZipArchive\File2.txt" "This is test file 2." -PassThru;
        $File3 = Setup -File "NewZipArchive\File3.txt" "This is test file 3." -PassThru;
        
        (Get-Item -Path "TestDrive:\NewZipArchive").GetType() | Should Be System.IO.DirectoryInfo;
        (Get-ChildItem -Path "TestDrive:\NewZipArchive").Count | Should Be 3;
	}

    It "Creates 'NewZipArchive.zip'" {
        $NewZipArchive = New-ZipArchive -Path "TestDrive:\NewZipArchive" -DestinationPath "$testDirectory\NewZipArchive.zip";
        (Get-Item -Path "TestDrive:\NewZipArchive.zip" -ErrorAction SilentlyContinue) | Should Not Be $null;
    }

    It "Fails to expand and overwrite an existing Zip Archive" {
        $zipArchiveItems = Expand-ZipArchive -Path "TestDrive:\NewZipArchive.zip" -DestinationPath "$testDirectory\NewZipArchive" -WarningAction SilentlyContinue;
        $zipArchiveItems | Should Be $null;
    }

    It "Expands and overwrites an existing Zip Archive Item" {
        $zipArchiveItems = Expand-ZipArchive -Path "TestDrive:\NewZipArchive.zip" -DestinationPath "$testDirectory\NewZipArchive" -Force;
        #$zipArchiveItem = Expand-ZipArchiveItem -InputObject ([ref] $zipArchiveItems[0]) -DestinationPath "$testDirectory\NewZipArchive" -Force;
        $zipArchiveItems | Should Not Be $null;
    }

    It "Expands Zip Archive to a new directory" {
        ## TODO: Should this not create the target directory, or at least fail if it doesn't exist?
        Setup -Dir "NewZipArchive2";
        Expand-ZipArchive -Path "TestDrive:\NewZipArchive.zip" -DestinationPath "$testDirectory\NewZipArchive2" | Should Not Be Null;
    }
}

Describe "Expand-ZipArchiveItem" {

    ## Setup test drive
    $testDirectory = (New-TestDrive -PassThru).Root;

    It "Creates the test directory" {

        Setup -Dir "NewZipArchive" -PassThru | Should Exist;
        $File1 = Setup -File "NewZipArchive\File1.txt" "This is test file 1." -PassThru;
        $File2 = Setup -File "NewZipArchive\File2.txt" "This is test file 2." -PassThru;
        $File3 = Setup -File "NewZipArchive\File3.txt" "This is test file 3." -PassThru;
        
        (Get-Item -Path "TestDrive:\NewZipArchive").GetType() | Should Be System.IO.DirectoryInfo;
        (Get-ChildItem -Path "TestDrive:\NewZipArchive").Count | Should Be 3;
	}

    It "Creates 'NewZipArchive.zip'" {
        $NewZipArchive = New-ZipArchive -Path "TestDrive:\NewZipArchive" -DestinationPath "$testDirectory\NewZipArchive.zip";
        (Get-Item -Path "TestDrive:\NewZipArchive.zip" -ErrorAction SilentlyContinue) | Should Not Be $null;
    }

    It "Expands and overwrite existing Zip Archive Items" {
        $zipArchiveItems = Get-ZipArchiveEntry -Path "TestDrive:\NewZipArchive.zip" | Expand-ZipArchiveItem -DestinationPath "$testDirectory\NewZipArchive" -Force;
        $zipArchiveItems | Should Not Be $null;
    }

    It "Expands and does not overwrite existing Zip Archive Items" {
        $zipArchiveItems = Get-ZipArchiveEntry -Path "TestDrive:\NewZipArchive.zip" | Expand-ZipArchiveItem -DestinationPath "$testDirectory\NewZipArchive" -WarningAction SilentlyContinue;
        $zipArchiveItems | Should Be $null;
    }

}

Describe "Add-ZipArchiveItem" {
    
    ## Setup test drive
    $testDirectory = (New-TestDrive -PassThru).Root;

    It "Creates the test directory" {
        Setup -Dir "NewZipArchive" -PassThru | Should Exist;
        $File1 = Setup -File "NewZipArchive\File1.txt" "This is test file 1." -PassThru;
        $File2 = Setup -File "NewZipArchive\File2.txt" "This is test file 2." -PassThru;
        $File3 = Setup -File "NewZipArchive\File3.txt" "This is test file 3." -PassThru;
        
        (Get-Item -Path "TestDrive:\NewZipArchive").GetType() | Should Be System.IO.DirectoryInfo;
        (Get-ChildItem -Path "TestDrive:\NewZipArchive").Count | Should Be 3;
	}

    It "Creates 'NewZipArchive.zip'" {
        $NewZipArchive = New-ZipArchive -Path "TestDrive:\NewZipArchive" -DestinationPath "$testDirectory\NewZipArchive.zip";
        (Get-Item -Path "TestDrive:\NewZipArchive.zip" -ErrorAction SilentlyContinue) | Should Not Be $null;
    }

    It "Adds a file to an existing Zip Archive" {
        $File4 = Setup -File "NewZipArchive\File4.txt" "This is test file 4." -PassThru;
        Add-ZipArchiveItem -Path "TestDrive:\NewZipArchive\File4.txt" -DestinationPath "$testDirectory\NewZipArchive.zip";
        (Get-Item -Path "TestDrive:\NewZipArchive.zip").Length | Should Be 478;
    }

    It "Checks the Zip Archive contents" {
        (Get-ZipArchiveEntry -Path "TestDrive:\NewZipArchive.zip").Count | Should Be 4;
    }
   
    It "Fails to add a file to an existing Zip Archive" {
        $File4 = Setup -File "NewZipArchive\File4.txt" "This is a different test file 4." -PassThru;
        Add-ZipArchiveItem -Path "TestDrive:\NewZipArchive\File4.txt" -DestinationPath "$testDirectory\NewZipArchive.zip" -WarningAction SilentlyContinue;
        (Get-Item -Path "TestDrive:\NewZipArchive.zip").Length | Should Be 478;
    }

    It "Overwrites an existing Zip Archive Item" {
        Add-ZipArchiveItem -Path "TestDrive:\NewZipArchive\File4.txt" -DestinationPath "$testDirectory\NewZipArchive.zip" -Force;
        (Get-Item -Path "TestDrive:\NewZipArchive.zip").Length | Should Be 490;
    }

    It "Adds files to an existing Zip Archive by string array" {
        $File5 = Setup -File "NewZipArchive\File5.txt" "This is test file 5." -PassThru;
        $File6 = Setup -File "NewZipArchive\File6.txt" "This is test file 6." -PassThru;

        $files = Get-ChildItem -Path "TestDrive:\NewZipArchive\File*" -Recurse;
        Add-ZipArchiveItem -Path $files -DestinationPath "$testDirectory\NewZipArchive.zip" -WarningAction SilentlyContinue;
        (Get-Item -Path "TestDrive:\NewZipArchive.zip").Length | Should Be 718;
    }

    It "Checks the Zip Archive contents" {
        (Get-ZipArchiveEntry -Path "TestDrive:\NewZipArchive.zip").Count | Should Be 6;
    }
}
