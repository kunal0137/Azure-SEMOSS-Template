Write-Host 'Please allow several minutes for the install to complete. '


# Install Google Chrome x64 on 64-Bit systems? $True or $False
$Installx64 = $True

# Define the temporary location to cache the installer.
$TempDirectory = "$ENV:Temp\Chrome"

# Run the script silently, $True or $False
$RunScriptSilent = $True

# Set the system architecture as a value.
$OSArchitecture = (Get-WmiObject Win32_OperatingSystem).OSArchitecture

# Exit if the script was not run with Administrator priveleges
$User = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
if (-not $User.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
	Write-Host 'Please run again with Administrator privileges.' -ForegroundColor Red
    if ($RunScriptSilent -NE $True){
        Read-Host 'Press [Enter] to exit'
    }
    exit
}


Function Download-Chrome {
    Write-Host 'Downloading Google Chrome... ' -NoNewLine

    # Test internet connection
    if (Test-Connection google.com -Count 3 -Quiet) {
		if ($OSArchitecture -eq "64-Bit" -and $Installx64 -eq $True){
			$Link = 'http://dl.google.com/edgedl/chrome/install/GoogleChromeStandaloneEnterprise64.msi'
		} ELSE {
			$Link = 'http://dl.google.com/edgedl/chrome/install/GoogleChromeStandaloneEnterprise.msi'
		}
    

        # Download the installer from Google
        try {
	        New-Item -ItemType Directory "$TempDirectory" -Force | Out-Null
	        (New-Object System.Net.WebClient).DownloadFile($Link, "$TempDirectory\Chrome.msi")
            Write-Host 'success!' -ForegroundColor Green
        } catch {
	        Write-Host 'failed. There was a problem with the download.' -ForegroundColor Red
            if ($RunScriptSilent -NE $True){
                Read-Host 'Press [Enter] to exit'
            }
	        exit
        }
    } else {
        Write-Host "failed. Unable to connect to Google's servers." -ForegroundColor Red
        if ($RunScriptSilent -NE $True){
            Read-Host 'Press [Enter] to exit'
        }
	    exit
    }
}

Function Install-Chrome {
    Write-Host 'Installing Chrome... ' -NoNewline

    # Install Chrome
    $ChromeMSI = """$TempDirectory\Chrome.msi"""
	$ExitCode = (Start-Process -filepath msiexec -argumentlist "/i $ChromeMSI /qn /norestart" -Wait -PassThru).ExitCode
    
    if ($ExitCode -eq 0) {
        Write-Host 'success!' -ForegroundColor Green
    } else {
        Write-Host "failed. There was a problem installing Google Chrome. MsiExec returned exit code $ExitCode." -ForegroundColor Red
        Clean-Up
        if ($RunScriptSilent -NE $True){
            Read-Host 'Press [Enter] to exit'
        }
	    exit
    }
}

Function Clean-Up {
    Write-Host 'Removing Chrome installer... ' -NoNewline

    try {
        # Remove the installer
        Remove-Item "$TempDirectory\Chrome.msi" -ErrorAction Stop
        Write-Host 'success!' -ForegroundColor Green
    } catch {
        Write-Host "failed. You will have to remove the installer yourself from $TempDirectory\." -ForegroundColor Yellow
    }
}

Download-Chrome
Install-Chrome
Clean-Up

if ($RunScriptSilent -NE $True){
    Read-Host 'Install complete! Press [Enter] to exit'
}

#Install SEMOSS

#Set Directory to download SEMOSS
$TempDirectory = "$ENV:Temp\SEMOSS"
Function Download-SEMOSS {


    Write-Host 'Downloading SEMOSS... ' -NoNewLine

    # Test internet connection
    if (Test-Connection semoss.org -Count 3 -Quiet) {
			$Link = 'http://semoss.org/download/SEMOSS_v3.2_x64.zip'
		 
    

        # Download the installer from Google
        try {
	        New-Item -ItemType Directory "$TempDirectory" -Force | Out-Null
	        (New-Object System.Net.WebClient).DownloadFile($Link, "$TempDirectory\SEMOSS.zip")
            Write-Host 'success!' -ForegroundColor Green
        } catch {
	        Write-Host 'failed. There was a problem with the download.' -ForegroundColor Red
            if ($RunScriptSilent -NE $True){
                Read-Host 'Press [Enter] to exit'
            }
	        exit
        }
    } else {
        Write-Host "failed. Unable to connect to SEMOSS's servers." -ForegroundColor Red
        if ($RunScriptSilent -NE $True){
            Read-Host 'Press [Enter] to exit'
        }
	    exit
    }
}


Function Install-SEMOSS {
    Write-Host 'Installing SEMOSS... ' -NoNewline

    # Install SEMOSS
    $SEMOSSzip = """$TempDirectory\SEMOSS.zip"""

    
    $w = $null

	Expand-Archive $TempDirectory\SEMOSS.zip -DestinationPath c:\ -WarningVariable w

    if($w.Count -gt 0){
    Write-Host "failed. Unable unzip files." -ForegroundColor Red 
}
else{
Write-Host 'success!' -ForegroundColor Green
}
}

Function Clean-Up {
    Write-Host 'Removing SEMOSS install files... ' -NoNewline

    try {
        # Remove the installer
        Remove-Item "$TempDirectory\SEMOSS.zip" -ErrorAction Stop
        Write-Host 'success!' -ForegroundColor Green
    } catch {
        Write-Host "failed. You will have to remove the installer yourself from $TempDirectory\." -ForegroundColor Yellow
    }
}
Download-SEMOSS

Install-SEMOSS
Clean-up

Write-Host 'Opening firewall ports'
netsh advfirewall firewall add rule name="Open Port 80" dir=in action=allow protocol=TCP localport=80
netsh advfirewall firewall add rule name="Open Port 5355" dir=in action=allow protocol=TCP localport=5355

Write-Host 'Editing start-up scripts'

$Act = New-ScheduledTaskAction -Execute "C:\SEMOSS_v3.2_x64\startSEMOSS.bat" -WorkingDirectory C:\SEMOSS_v3.2_x64
$T = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask StartSemossStartup -Action $Act -Trigger $T -Principal $principal



start powershell { cd c:\SEMOSS_v3.2_x64; 'n'|.\startSEMOSS.bat}

