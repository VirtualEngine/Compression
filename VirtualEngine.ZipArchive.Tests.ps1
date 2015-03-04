## Load the required .NET assemblies
Add-Type -AssemblyName "System.IO.Compression";
Add-Type -AssemblyName "System.IO.Compression.FileSystem";

$here = Split-Path -Parent $MyInvocation.MyCommand.Path;
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".");
. "$here\$sut"

Describe "New-ZipArchive" {

    It "Creates the test directory" {
        New-Item -Path 'TestDrive:\NewZipArchive' -ItemType Directory | Should Exist;
        $File1 = Set-Content -Path 'TestDrive:\NewZipArchive\File1.txt' -Value 'This is test file 1.';
        $File2 = Set-Content -Path 'TestDrive:\NewZipArchive\File2.txt' -Value 'This is test file 2.';
        $File3 = Set-Content -Path 'TestDrive:\NewZipArchive\File3.txt' -Value 'This is test file 3.';       
        (Get-Item -Path 'TestDrive:\NewZipArchive').GetType() | Should Be System.IO.DirectoryInfo;
        (Get-ChildItem -Path 'TestDrive:\NewZipArchive').Count | Should Be 3;  
	}

    It "Creates 'NewZipArchive.zip' from directory" {
        $NewZipArchive = New-ZipArchive -Path "TestDrive:\NewZipArchive" -DestinationPath 'TestDrive:\NewZipArchive.zip';
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip' -ErrorAction SilentlyContinue) | Should Not Be $null;
    }
    
    It "Checks 'NewZipArchive.zip' size" {
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip').Length | Should Be 370;
    }

    It "Creates 'NewZipArchive.zip' from string array" {
        $files = Get-ChildItem -Path 'TestDrive:\NewZipArchive\*' -Recurse;
        $NewZipArchive = New-ZipArchive -Path $files.FullName -DestinationPath 'TestDrive:\NewZipArchive.zip';
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip' -ErrorAction SilentlyContinue) | Should Not Be $null;
    }

    It "Checks 'NewZipArchive.zip' size" {
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip').Length | Should Be 370;
    }
}

Describe "Get-ZipArchiveItem" {

    It "Creates the test directory" {
        New-Item -Path 'TestDrive:\NewZipArchive' -ItemType Directory | Should Exist;
        $File1 = Set-Content -Path 'TestDrive:\NewZipArchive\File1.txt' -Value 'This is test file 1.';
        $File2 = Set-Content -Path 'TestDrive:\NewZipArchive\File2.txt' -Value 'This is test file 2.';
        $File3 = Set-Content -Path 'TestDrive:\NewZipArchive\File3.txt' -Value 'This is test file 3.';   
        (Get-Item -Path 'TestDrive:\NewZipArchive').GetType() | Should Be System.IO.DirectoryInfo;
        (Get-ChildItem -Path 'TestDrive:\NewZipArchive').Count | Should Be 3;
	}

    It "Creates 'NewZipArchive.zip'" {
        $NewZipArchive = New-ZipArchive -Path 'TestDrive:\NewZipArchive' -DestinationPath 'TestDrive:\NewZipArchive.zip';
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip' -ErrorAction SilentlyContinue) | Should Not Be $null;
    }
    
    It "Ensures the Zip Archive contents" {
        $zipArchiveItems = Get-ZipArchiveItem -Path 'TestDrive:\NewZipArchive.zip';
        $zipArchiveItems.Count | Should Be 3;
    }
}

Describe "Expand-ZipArchive" {

    It "Creates the test directory" {
        New-Item -Path 'TestDrive:\NewZipArchive' -ItemType Directory | Should Exist;
        $File1 = Set-Content -Path 'TestDrive:\NewZipArchive\File1.txt' -Value 'This is test file 1.';
        $File2 = Set-Content -Path 'TestDrive:\NewZipArchive\File2.txt' -Value 'This is test file 2.';
        $File3 = Set-Content -Path 'TestDrive:\NewZipArchive\File3.txt' -Value 'This is test file 3.';   
        (Get-Item -Path 'TestDrive:\NewZipArchive').GetType() | Should Be System.IO.DirectoryInfo;
        (Get-ChildItem -Path 'TestDrive:\NewZipArchive').Count | Should Be 3;
	}

    It "Creates 'NewZipArchive.zip'" {
        $NewZipArchive = New-ZipArchive -Path 'TestDrive:\NewZipArchive' -DestinationPath 'TestDrive:\NewZipArchive.zip';
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip' -ErrorAction SilentlyContinue) | Should Not Be $null;
    }

    It "Fails to expand and overwrite an existing Zip Archive" {
        $zipArchiveItems = Expand-ZipArchive -Path 'TestDrive:\NewZipArchive.zip' -DestinationPath 'TestDrive:\NewZipArchive' -WarningAction SilentlyContinue;
        $zipArchiveItems.Count | Should Be 0;
        (Get-ChildItem -Path 'TestDrive:\NewZipArchive').Count | Should Be 3;
    }

    It "Expands and overwrites an existing Zip Archive Item" {
        $zipArchiveItems = Expand-ZipArchive -Path 'TestDrive:\NewZipArchive.zip' -DestinationPath 'TestDrive:\NewZipArchive' -Force;
        $zipArchiveItems.Count | Should Be 3;
        (Get-ChildItem -Path 'TestDrive:\NewZipArchive').Count | Should Be 3;
    }

    It "Expands Zip Archive to a new directory" {
        $zipArchiveItems = Expand-ZipArchive -Path 'TestDrive:\NewZipArchive.zip' -DestinationPath 'TestDrive:\NewZipArchive2';
        $zipArchiveItems.Count | Should Be 3;
        (Get-ChildItem -Path 'TestDrive:\NewZipArchive2').Count | Should Be 3;
    }
}

Describe "Expand-ZipArchiveItem" {

    ## Resolve TestDrive:\ location for underlying .Net methods
    $testDrive = Get-PSDrive TestDrive | Select -ExpandProperty Root;

    It "Creates the test directory" {
        New-Item -Path 'TestDrive:\NewZipArchive' -ItemType Directory | Should Exist;
        $File1 = Set-Content -Path 'TestDrive:\NewZipArchive\File1.txt' -Value 'This is test file 1.';
        $File2 = Set-Content -Path 'TestDrive:\NewZipArchive\File2.txt' -Value 'This is test file 2.';
        $File3 = Set-Content -Path 'TestDrive:\NewZipArchive\File3.txt' -Value 'This is test file 3.';   
        (Get-Item -Path 'TestDrive:\NewZipArchive').GetType() | Should Be System.IO.DirectoryInfo;
        (Get-ChildItem -Path 'TestDrive:\NewZipArchive').Count | Should Be 3;
	}

    It "Creates 'NewZipArchive.zip'" {
        $NewZipArchive = New-ZipArchive -Path 'TestDrive:\NewZipArchive' -DestinationPath "$testDrive\NewZipArchive.zip";
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip' -ErrorAction SilentlyContinue) | Should Not Be $null;
    }

    It "Expands and overwrite existing Zip Archive Items" {
        $zipArchiveItems = Get-ZipArchiveItem -Path 'TestDrive:\NewZipArchive.zip' | Expand-ZipArchiveItem -DestinationPath "$testDrive\NewZipArchive" -Force;
        $zipArchiveItems.Count | Should Be 3;
    }

    It "Expands and does not overwrite existing Zip Archive Items" {
        $zipArchiveItems = Get-ZipArchiveItem -Path 'TestDrive:\NewZipArchive.zip' | Expand-ZipArchiveItem -DestinationPath "$testDrive\NewZipArchive" -WarningAction SilentlyContinue;
        $zipArchiveItems | Should Be $null;
    }

}

Describe "Add-ZipArchiveItem" {
    
    It "Creates the test directory" {
        New-Item -Path 'TestDrive:\NewZipArchive' -ItemType Directory | Should Exist;
        $File1 = Set-Content -Path 'TestDrive:\NewZipArchive\File1.txt' -Value 'This is test file 1.';
        $File2 = Set-Content -Path 'TestDrive:\NewZipArchive\File2.txt' -Value 'This is test file 2.';
        $File3 = Set-Content -Path 'TestDrive:\NewZipArchive\File3.txt' -Value 'This is test file 3.';
        (Get-Item -Path 'TestDrive:\NewZipArchive').GetType() | Should Be System.IO.DirectoryInfo;
        (Get-ChildItem -Path 'TestDrive:\NewZipArchive').Count | Should Be 3;
	}

    It "Creates 'NewZipArchive.zip'" {
        $NewZipArchive = New-ZipArchive -Path 'TestDrive:\NewZipArchive' -DestinationPath 'TestDrive:\NewZipArchive.zip';
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip' -ErrorAction SilentlyContinue) | Should Not Be $null;
    }

    It "Adds a file to an existing Zip Archive" {
        $File4 = Set-Content -Path 'TestDrive:\NewZipArchive\File4.txt' -Value 'This is test file 4.';
        Add-ZipArchiveItem -Path 'TestDrive:\NewZipArchive\File4.txt' -DestinationPath 'TestDrive:\NewZipArchive.zip';
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip').Length | Should Be 486;
    }

    It "Checks the Zip Archive contents" {
        (Get-ZipArchiveItem -Path 'TestDrive:\NewZipArchive.zip').Count | Should Be 4;
    }
   
    It "Fails to add a file to an existing Zip Archive" {
        $File4 = Set-Content -Path 'TestDrive:\NewZipArchive\File4.txt' -Value 'This is a different test file 4.';
        Add-ZipArchiveItem -Path 'TestDrive:\NewZipArchive\File4.txt' -DestinationPath 'TestDrive:\NewZipArchive.zip' -WarningAction SilentlyContinue;
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip').Length | Should Be 486;
    }

    It "Overwrites an existing Zip Archive Item" {
        Add-ZipArchiveItem -Path 'TestDrive:\NewZipArchive\File4.txt' -DestinationPath 'TestDrive:\NewZipArchive.zip' -Force;
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip').Length | Should Be 498;
    }

    It "Adds files to an existing Zip Archive by string array" {
        $File5 = Set-Content -Path 'TestDrive:\NewZipArchive\File5.txt' -Value 'This is test file 5.';
        $File6 = Set-Content -Path 'TestDrive:\NewZipArchive\File6.txt' -Value 'This is test file 6.';
        $files = Get-ChildItem -Path 'TestDrive:\NewZipArchive\File*' -Recurse;
        Add-ZipArchiveItem -Path $files -DestinationPath 'TestDrive:\NewZipArchive.zip' -WarningAction SilentlyContinue;
        (Get-Item -Path 'TestDrive:\NewZipArchive.zip').Length | Should Be 730;
    }

    It "Checks the Zip Archive contents" {
        (Get-ZipArchiveItem -Path 'TestDrive:\NewZipArchive.zip').Count | Should Be 6;
    }
}
