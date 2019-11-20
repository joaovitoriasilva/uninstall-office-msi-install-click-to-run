<#

.SYNOPSIS
This script will detect Office Versions (2003 and up) installed in the machine where is being runned and will remove them. After remove them, it will install the Office 365 ProPlus click-to-run version.
If the script detects a version of Visio and/or Project installed, it will also install the click-to-run version.

.DESCRIPTION
The script will uninstall Office versions and install the click-to-run versions. Will output to a log file stored in %temp%\Office365Deploy.log info from the execution of the file.

.Parameter installVisio365
Variable to define if Visio is found by the script, will it install the Visio 365 Version. By default $false.

.Parameter installProject365
Variable to define if Project is found by the script, will it install the Project 365 Version. By default $false.

.Notes
Script created by João Vitória Silva: https://www.linkedin.com/in/joao-v-silva/
Version 1.4 - Office365ProPlus_Deploy_v1.4-JVS-29-03-2019
Script changes:	1.0 - Initial version
                1.1 - Added reg keys to remove first run setup
				1.2 - Added support for french and portuguese OS LP
				1.3 - Added support for uninstalling previously installed Office 365 ProPlus versions
                      Fixed error when calling Office 365 ProPlus installer with spaces in the path
                1.4 - Added detection for solo installations of Office 365 Project and/or Visio clients

.EXAMPLE
.\Office365ProPlusDeploy.ps1
.\Office365ProPlusDeploy.ps1 -installProject365 $true
.\Office365ProPlusDeploy.ps1 -installVisio365 $true
.\Office365ProPlusDeploy.ps1 -installVisio365 $true -installProject365 $true

#>

Param ([switch]$installProject365 = $false,
[switch]$installVisio365 = $false)

# Functions Start
# Stores script run info to log file
Function Write-Log{
   Param ([string]$logString,
        [string]$color)

   Write-Host $logstring -ForegroundColor $color
   try{
       Add-Content $logfile -value $logstring -ErrorAction Stop
   }catch{
       Write-Host "Unable to write to log file!" -ForegroundColor "Red"
   }
}

# Will search for LP installed for that version of Office
Function Get-OfficeLPsInstalled{
    Param ([string]$path)
   
    if (Test-Path $path){
        foreach ($lp in $lps.GetEnumerator()){ 
            $aux = $lp.Value
            try {
                # Query registry for LP
                Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $lp.Name -ErrorAction Stop | Out-Null
                Write-Log -logString "Office $version $officeArch LanguagePack $aux exists, uninstalling..." -Color "White"
                # Call function to remove LP
                Remove-OfficeProduct -path $office -product "OMUI.$aux"
            }catch{
                Write-Log -logString "Language pack $aux not found!" -Color "Yellow"
            }
        }
    }
}

# Remove Office MSI Product
Function Remove-OfficeProduct{
    Param (
        [string]$path,
        [string]$product
    )

    # Checking for Project Pro installation
    if ($product -eq "PrjPro"){
        $script:isProjectProInstalled = $true
        Write-Log -logString "Uninstalling Project Pro $version $officeArch" -color "White"
    }else{
        # Checking for Project Std installation
        if($product -eq "PrjStd"){
            $script:isProjectStdInstalled = $true
            Write-Log -logString "Uninstalling Project Std $version $officeArch" -color "White"
        }else{
            # Checking for Sharepoint Designer installation
            if ($product -eq "SharePointDesigner"){
            #    $script:isShptDesInstalled = $true
                Write-Log -logString "Uninstalling SharePoint Designer $version $officeArch" -color "White"
            }else{
                # Checking for Visio Pro installation
                if ($product -eq "VISPRO"){
                    $script:isVisioProInstalled = $true
                    Write-Log -logString "Uninstalling Visio Pro $version $officeArch" -color "White"
                }else{
                    # Checking for Visio Std installation
                    if($product -eq "VisStd"){
                        $script:isVisioStdInstalled = $true
                        Write-Log -logString "Uninstalling Visio Std $version $officeArch" -color "White"
                    }else{
                        # Checking for Lync installation
                        if(($product -eq "LYNC") -or ($product -eq "LYNCENTRY")){
                            Write-Log -logString "Uninstalling Lync/Skype for Business $version $officeArch" -color "White"
                        }else{
                            # Checking for Language Pack installation
                            if(!($product -like "OMUI*")){
                                Write-Log -logString "Uninstalling Office $version $officeArch" -color "White"
                            }
                        }
                    }
                }
            }
        }
    }

    $arguments = "/uninstall $product /config `"$uninstallXML`""
    try {
        Start-Process -FilePath "$path" -ArgumentList $arguments -Wait -NoNewWindow -ErrorAction Stop
    }catch{
        Write-Log -logString "It wasn't possible to uninstall $product product!" -color "Red"
        Write-Log -logString "LanguagePack? It may be the Office language and not a Language Pack..." -color "Red"
    }
}

# Remove Office 2003
Function Remove-Office2003{
    $office2003x86x86Versions=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90110816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90130816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90CA0816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90E30816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{903A0816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{903B0816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90510816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90520816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90530816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90330816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{901F0409-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{901E0816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90230816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90FF0816-6000-11D3-8CFE-0150048383C9}")
    $office2003x86x64Versions=@("HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90110816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90130816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90CA0816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90E30816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{903A0816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{903B0816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90510816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90520816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90530816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90330816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{901F0409-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{901E0816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90230816-6000-11D3-8CFE-0150048383C9}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90FF0816-6000-11D3-8CFE-0150048383C9}")
    # Setting array with Office 2007 product ids
    $office2003ProductID=@("PRO",
                            "STANDARD",
                            "BASIC",
                            "SmallBusiness",
                            "PROPLUS",
                            "PrjStd",
                            "PrjPro",
                            "VisPro",
                            "VisView",
                            "VisStd",
                            "PERSONAL",
                            "PTK",
                            "OfficeMUI",
                            "OfficeMUI",
                            "LIP")
    if (($osArch -eq "x86") -and ($officeArch -eq "x86")){
        # For Loop to iterate correct array and Product Code array
        for($i=0
        $i -lt $office2003ProductID.Count
        $i++){
            if (Test-Path $office2003x86x86Versions[$i]){
                # Uninstalling Office Product
                $key=$office2003x86x86Versions[$i].split("\")
                $keyAux=$key[-1]
                $arguments = "/x $keyAux /qb /norestart "
                try {
                    Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -NoNewWindow -ErrorAction Stop
                }catch{
                    Write-Log -logString "It wasn't possible to uninstall $product product!" -ForegroundColor "Red"
                }
            }
        }
    }else{
        if (($osArch -eq "x64") -and ($officeArch -eq "x86")){
            # For Loop to iterate correct array and Product Code array
            for($i=0
            $i -lt $office2003ProductID.Count
            $i++){
                if (Test-Path $office2003x86x64Versions[$i]){
                    # Uninstalling Office Product
                    $key=$office2003x86x64Versions[$i].split("\")
                    $keyAux=$key[-1]
                    $arguments = "/x $keyAux /qb /norestart "
                    try {
                        Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -NoNewWindow -ErrorAction Stop
                    }catch{
                        Write-Log -logString "It wasn't possible to uninstall $product product!" -ForegroundColor "Red"
                    } 
                }
            }
        }
    }
}

# Remove Office 2007 and up
Function Remove-Office2007andUp{
    Param (
        [string]$lpsPathWOW,
        [string]$lpsPath
    )

    # https://support.microsoft.com/en-us/help/928516/description-of-product-code-guids-in-2007-office-suites-and-programs
    if ($version -eq 2007){
        $officex64x64Versions=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0011-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0012-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0013-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0014-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-002E-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-002F-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0030-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0031-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0033-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0035-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-00CA-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0017-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-003A-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-003B-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0051-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0052-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0053-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0074-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-004B-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-012D-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-012B-0000-1000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-001F-0000-1000-0000000FF1CE}")
        $officex86x64Versions=@("HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0011-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0012-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0013-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0014-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-002E-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-002F-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0030-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0031-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0033-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0035-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-00CA-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0017-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-003A-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-003B-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0051-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0052-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0053-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0074-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-004B-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-012D-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-012B-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-001F-0000-0000-0000000FF1CE}")
        $officex86x86Versions=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0011-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0012-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0013-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0014-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-002E-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-002F-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0030-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0031-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0033-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0035-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-00CA-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0017-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-003A-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-003B-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0051-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0052-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0053-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-0074-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-004B-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-012D-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-012B-0000-0000-0000000FF1CE}",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-001F-0000-0000-0000000FF1CE}")
    }else{
        if ($version -eq 2010){
            $officex64x64Versions=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0011-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0012-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0013-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0014-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-002E-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-002F-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0030-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0031-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0033-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0035-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-00CA-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0017-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-003A-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-003B-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0051-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0052-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0053-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0074-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-004B-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-012D-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-012B-0000-1000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-001F-0000-1000-0000000FF1CE}")
            $officex86x64Versions=@("HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0011-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0012-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0013-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0014-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-002E-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-002F-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0030-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0031-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0033-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0035-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-00CA-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0017-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-003A-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-003B-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0051-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0052-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0053-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0074-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-004B-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-012D-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-012B-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-001F-0000-0000-0000000FF1CE}")
            $officex86x86Versions=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0011-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0012-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0013-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0014-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-002E-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-002F-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0030-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0031-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0033-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0035-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-00CA-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0017-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-003A-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-003B-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0051-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0052-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0053-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-0074-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-004B-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-012D-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-012B-0000-0000-0000000FF1CE}",
                                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90140000-001F-0000-0000-0000000FF1CE}")
        }else{
            if ($version -eq 2013){
                $officex64x64Versions=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0011-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0012-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0013-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0014-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-002E-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-002F-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0030-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0031-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0033-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0035-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-00CA-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0017-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-003A-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-003B-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0051-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0052-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0053-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0074-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-004B-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-012D-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-012B-0000-1000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-001F-0000-1000-0000000FF1CE}")
                $officex86x64Versions=@("HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0011-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0012-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0013-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0014-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-002E-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-002F-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0030-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0031-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0033-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0035-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-00CA-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0017-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-003A-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-003B-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0051-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0052-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0053-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0074-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-004B-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-012D-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-012B-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-001F-0000-0000-0000000FF1CE}")
                $officex86x86Versions=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0011-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0012-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0013-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0014-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-002E-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-002F-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0030-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0031-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0033-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0035-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-00CA-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0017-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-003A-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-003B-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0051-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0052-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0053-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-0074-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-004B-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-012D-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-012B-0000-0000-0000000FF1CE}",
                                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90150000-001F-0000-0000-0000000FF1CE}")
            }else{
                if ($version -eq 2016){
                    $officex64x64Versions=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0011-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0012-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0013-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0014-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-002E-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-002F-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0030-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0031-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0033-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0035-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-00CA-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0017-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-003A-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-003B-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0051-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0052-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0053-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0074-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-004B-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-012D-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-012B-0000-1000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-001F-0000-1000-0000000FF1CE}")
                    $officex86x64Versions=@("HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0011-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0012-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0013-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0014-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-002E-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-002F-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0030-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0031-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0033-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0035-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-00CA-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0017-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-003A-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-003B-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0051-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0052-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0053-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0074-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-004B-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-012D-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-012B-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-001F-0000-0000-0000000FF1CE}")
                    $officex86x86Versions=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0011-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0012-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0013-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0014-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-002E-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-002F-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0030-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0031-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0033-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0035-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-00CA-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0017-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-003A-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-003B-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0051-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0052-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0053-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0074-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-004B-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-012D-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-012B-0000-0000-0000000FF1CE}",
                                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-001F-0000-0000-0000000FF1CE}")
                }
            }
        }
    }
    # Setting array with Office 2007 product ids
    $officeProductID=@("PROPLUS",
                            "STANDARD",
                            "BASIC",
                            "PRO",
                            "Ultimate",
                            "HomeAndStudent",
                            "ENTERPRISE",
                            "ProfessionalHybrid",
                            "Personal",
                            "ProfessionalHybrid",
                            "SmallBusiness",
                            "SharePointDesigner",
                            "PrjStd",
                            "PrjPro",
                            "VISPRO",
                            "VisView",
                            "VisStd",
                            "Starter",
                            "PROOFKIT",
                            "LYNCENTRY",
                            "LYNC",
                            "PROOF")
    if (($osArch -eq "x86") -and ($officeArch -eq "x86")){
        # Checking for installed language packs
        Get-OfficeLPsInstalled -path $lpsPath
        # For Loop to iterate correct array and Product Code array
        for($i=0
        $i -lt $officeProductID.Count
        $i++){
            if (Test-Path $officex86x86Versions[$i]){
                # Uninstalling Office Product
                Remove-OfficeProduct -path $office -product $officeProductID[$i]
            }
        }
    }else{
        if (($osArch -eq "x64") -and ($officeArch -eq "x86")){
            # Checking for installed language packs
            Get-OfficeLPsInstalled -path $lpsPathWOW
            # For Loop to iterate correct array and Product Code array
            for($i=0
            $i -lt $officeProductID.Count
            $i++){
                if (Test-Path $officex86x64Versions[$i]){
                    # Uninstalling Office Product
                    Remove-OfficeProduct -path $office -product $officeProductID[$i]
                }
            }
        }else{
            if (($osArch -eq "x64") -and ($officeArch -eq "x64")){
                # Checking for installed language packs
                Get-OfficeLPsInstalled -path $lpsPath
                # For Loop to iterate correct array and Product Code array
                for($i=0
                $i -lt $officeProductID.Count
                $i++){
                    if (Test-Path $officex64x64Versions[$i]){
                        # Uninstalling Office Product
                        Remove-OfficeProduct -path $office -product $officeProductID[$i]
                    }
                }
            }
        }
    }
}

# Check for Office Versions and remove them
Function Remove-InstalledOfficeProducts{
    foreach($office in $officeVersions){
        if (Test-Path $office) {
            # Setting auxiliar variables
            $officeArch="x64"
            $version=2016
            # Getting Office arch installed and setting variable
            if ($office -like "*(x86)*") { 
                $officeArch="x86"
            }
            # Getting Office version installed and setting variable
            if ($office -like "*OFFICE15*"){
                $version=2013
            }else{
                if ($office -like "*OFFICE14*"){
                    $version=2010
                }else{
                    if ($office -like "*OFFICE12*"){
                        $version=2007
                    }else{
                        if ($office -like "*OFFICE11*"){
                            $version=2003
                        }else{
							if (($office -like "*Office16\WINWORD.EXE") -or ($office -like "*Office15\WINWORD.EXE")){
								$version=365
							}else{
                                if (($office -like "*Office16\WINPROJ.EXE") -or ($office -like "*Office15\WINPROJ.EXE")){
                                    $version="365Project"
                                }else{
                                    if (($office -like "*Office16\VISIO.EXE") -or ($office -like "*Office15\VISIO.EXE")){
                                        $version="365Visio"
                                    }
                                }
                            }
						}
                    }
                }
            }

            # Logging version to uninstall
            Write-Log -logString "Office $version $officeArch exists, uninstalling..." -Color "White"
            if ($version -eq 2003){
                Remove-Office2003
            }else{
                if ($version -eq 2007){
                    Remove-Office2007andUp -lpsPath "HKLM:\SOFTWARE\Microsoft\Office\12.0\Common\LanguageResources\InstalledUIs" -lpsPathWOW "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\12.0\Common\LanguageResources\InstalledUIs"
                }else{
                    if ($version -eq 2010){
                        Remove-Office2007andUp -lpsPath "HKLM:\SOFTWARE\Microsoft\Office\14.0\Common\LanguageResources\InstalledUIs" -lpsPathWOW "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\14.0\Common\LanguageResources\InstalledUIs"
                    }else{
                        if ($version -eq 2013){
                            Remove-Office2007andUp -lpsPath "HKLM:\SOFTWARE\Microsoft\Office\15.0\Common\LanguageResources\InstalledUIs" -lpsPathWOW "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Common\LanguageResources\InstalledUIs"
                        }else{
                            if ($version -eq 2016){
                                Remove-Office2007andUp -lpsPath "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\LanguageResources\InstalledUIs" -lpsPathWOW "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\LanguageResources\InstalledUIs"
                            }else{
								if ($version -eq 365){
									Write-Host ""
									Start-Sleep -Seconds 5
									Write-Log -logString "Uninstalling current version of Office 365 ProPlus..." -color "White"
									Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\UninstallO365_All.xml"
								}else{
                                    if ($version -eq "365Project"){
                                        Write-Host ""
                                        Start-Sleep -Seconds 5
                                        Write-Log -logString "Uninstalling current version of Office 365 ProPlus..." -color "White"
                                        $script:isProject365Installed = $true
                                        Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\UninstallO365Project.xml"
                                    }else{
                                        if ($version -eq "365Visio"){
                                            Write-Host ""
                                            Start-Sleep -Seconds 5
                                            Write-Log -logString "Uninstalling current version of Office 365 ProPlus..." -color "White"
                                            $script:isVisio365Installed = $true
                                            Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\UninstallO365Visio.xml"
                                        }
                                    }
                                }
							}
                        }
                    }
                }
            }
        }
    }
}

# Install Office 365 click-to-run product
Function Install-Office365Product{
    Param (
        [string]$path,
        [string]$xmlPath
    )

    $arguments = "/configure `"$xmlPath`""
    try{
        Start-Process -FilePath "$path" -ArgumentList "$arguments" -Wait -NoNewWindow -ErrorAction Stop
    }catch{
        Write-Log -logString "It wasn't possible to install the product!"
    }
}

# Based on the software previously installed on the workstation, will install the Office 365 click-to-run equivalent
Function Install-Office365OfficeProducts{
    Write-Host ""
    Start-Sleep -Seconds 5
    Write-Log -logString "Installing Office 365 ProPlus..." -color "White"
    # Installing Office 365 ProPlus
    Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\InstallO365.xml"
    #if ($isShptDesInstalled){
    #    Write-Log -logString "Installing SharePoint Designer 2013 click to run version" -color "White"
        # Installing SharePoint Designer if it was previously installed in the workstation
    #    Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\InstallShptDesigner.xml"
    #}
    if($installVisio365){
        if($isVisioProInstalled -or $isVisioStdInstalled -or $isVisio365Installed){
            Write-Log -logString "Installing Visio 365 click to run version" -color "White"
            # Installing Visio 365 if it was previously installed in the workstation
            Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\InstallVisio365.xml"
        }
    }else{
        if ($isVisioProInstalled){
            Write-Log -logString "Installing Visio Pro click to run version" -color "White"
            # Installing Visio Pro if it was previously installed in the workstation
            Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\InstallVisioPro.xml"
        }
        if ($isVisioStdInstalled){
            Write-Log -logString "Installing Visio Standard click to run version" -color "White"
            # Installing Visio Std if it was previously installed in the workstation
            Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\InstallVisioStd.xml"
        }
    }
    if($installProject365){
        if ($isProjectProInstalled -or $isProjectStdInstalled -or $isProject365Installed){
            Write-Log -logString "Installing Project 365 click to run version" -color "White"
            # Installing Project 365 if it was previously installed in the workstation
            Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\InstallProject365.xml"
        }
    }else{
        if ($isProjectProInstalled){
            Write-Log -logString "Installing Project Pro click to run version" -color "White"
            # Installing Project Pro if it was previously installed in the workstation
            Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\InstallProjectPro.xml"
        }
        if ($isProjectStdInstalled){
            Write-Log -logString "Installing Project Standard click to run version" -color "White"
            # Installing Project Std if it was previously installed in the workstation
            Install-Office365Product -path "$PSScriptRoot\setup.exe" -xmlPath "$PSScriptRoot\InstallProjectStd.xml"
        }
    }
    try{
        if(!(Test-Path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings")) {
            new-item -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings" -force
        }
        New-ItemProperty -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings" -Name "Count" -Value "1" -PropertyType DWORD -Force
        if(!(Test-Path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\FirstRun")) {
            new-item -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\FirstRun" -force
        }
        New-ItemProperty -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\FirstRun" -Name "BootedRTM" -Value "1" -PropertyType DWORD -Force
        New-ItemProperty -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\FirstRun" -Name "disablemovie" -Value "1" -PropertyType DWORD -Force
        if(!(Test-Path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\Common\General")) {
            new-item -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\Common\General" -force
        }
        New-ItemProperty -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\Common\General" -Name "shownfirstrunoptin" -Value "1" -PropertyType DWORD -Force
        New-ItemProperty -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\Common\General" -Name "ShownFileFmtPrompt" -Value "1" -PropertyType DWORD -Force
        if(!(Test-Path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\Common\PTWatson")) {
            new-item -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\Common\PTWatson" -force
        }
        New-ItemProperty -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\Common\PTWatson" -Name "PTWOptIn" -Value "1" -PropertyType DWORD -Force
        New-ItemProperty -path "HKLM:SOFTWARE\Wow6432Node\Microsoft\Office\16.0\User Settings\MyCustomUserSettings\Create\Software\Microsoft\Office\16.0\Common" -Name "qmenable" -Value "1" -PropertyType DWORD -Force
    }catch{
        Write-Log -logString "Unable to add reg keys. Will prompt for user config in first run." -color "Yellow"
    }
}
## Functions end

# Setting script version variable
$scriptVersion="Office365ProPlus_Deploy_v1.4-JVS-29-03-2019"

# Setting lof file variable
$logfile = "$env:TEMP\Office365Deploy.log"

# Setting Office installation path for uninstall process
$officeVersions=@("C:\Program Files\Microsoft Office\OFFICE11",
                    "C:\Program Files (x86)\Microsoft Office\OFFICE11",
                    "C:\Program Files\Microsoft Office\OFFICE11\WINWORD.EXE",
                    "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE12\Office Setup Controller\setup.exe",
                    "C:\Program Files\Common Files\Microsoft Shared\OFFICE12\Office Setup Controller\setup.exe",
                    "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\Office Setup Controller\setup.exe",
                    "C:\Program Files\Common Files\Microsoft Shared\OFFICE14\Office Setup Controller\setup.exe",
                    "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE15\Office Setup Controller\setup.exe",
                    "C:\Program Files\Common Files\Microsoft Shared\OFFICE15\Office Setup Controller\setup.exe",
                    "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.exe",
                    "C:\Program Files\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.exe",
                    "C:\Program Files (x86)\Microsoft Office\root\Office15\WINPROJ.EXE",
                    "C:\Program Files\Microsoft Office\root\Office15\WINPROJ.EXE",
                    "C:\Program Files (x86)\Microsoft Office\root\Office15\VISIO.EXE",
					"C:\Program Files\Microsoft Office\root\Office15\VISIO.EXE",
                    "C:\Program Files (x86)\Microsoft Office\root\Office16\WINPROJ.EXE",
                    "C:\Program Files\Microsoft Office\root\Office16\WINPROJ.EXE",
                    "C:\Program Files (x86)\Microsoft Office\root\Office16\VISIO.EXE",
                    "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE",
                    "C:\Program Files (x86)\Microsoft Office\root\Office15\WINWORD.EXE",
                    "C:\Program Files\Microsoft Office\root\Office15\WINWORD.EXE",
					"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
                    "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")

# Setting variable with language pack versions of Office
$lps = @{"1025"="ar-sa";"1026"="bg-bg";"2052"="zh-cn";"1028"="zh-tw";"1050"="hr-hr";"1029"="cs-cz";"1030"="da-dk";"1043"="nl-nl";
            "1061"="et-ee";"1035"="fi-fi";"1036"="fr-fr";"1031"="de-de";"1032"="el-gr";"1037"="he-il";"1081"="hi-in";"1038"="hu-hu";
            "1057"="id-id";"1040"="it-it";"1041"="ja-jp";"1087"="kk-kz";"1042"="ko-kr";"1062"="lv-lv";"1063"="lt-lt";"1086"="ms-my";
            "1044"="nb-no";"1045"="pl-pl";"1046"="pt-br";"2070"="pt-pt";"1048"="ro-ro";"1049"="ru-ru";"2074"="sr-latn-cs";"1051"="sk-sk";
            "1060"="sl-si";"3082"="es-es";"1053"="sv-se";"1054"="th-th";"1055"="tr-tr";"1058"="uk-ua";"1066"="vi-vn";"1033"="en-us"}

# Setting isProductInstalled variables for future comparison
#$installVisio365=$true
#$installProject365=$true
$isVisioProInstalled=$false
$isVisioStdInstalled=$false
$isVisio365Installed=$false
$isProjectProInstalled=$false
$isProjectStdInstalled=$false
$isProject365Installed=$false

# Checking PowerShell version. If version is lesser than 3, set PSScriptRoot variable
if ($PSVersionTable.PSVersion.Major -lt 3) {
    $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    Write-Log -logString "Set PSScriptRoot variable - $PSScriptRoot" -color "Yellow"
}
# Setting uninstall XML location
$uninstallXML="$PSScriptRoot\uninstallOfficeNonO365.xml"

# Checking Windows architecture version and saving it to variable
try{
    if (((Get-WmiObject Win32_OperatingSystem).OSArchitecture -eq "64-bit") -or ((Get-WmiObject Win32_OperatingSystem).OSArchitecture -eq "64 bits")){
        $osArch="x64"
    }else{
        $osArch="x86"
    }
}catch{
    Write-Log -logString "Unable to query OS architecture... Exiting!" -color "Red"
    exit 1
}

Write-Log -logString "Starting script $scriptVersion" -color "White"
Remove-InstalledOfficeProducts
Install-Office365OfficeProducts
Write-Log -logString "Ending script $scriptVersion" -color "White"