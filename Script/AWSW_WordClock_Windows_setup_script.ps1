# Automatic Arduino IDE setup for WordClock by AWSW
# DO NOT CHANGE ANYTHING FROM THIS LINE ON ! # # DO NOT CHANGE ANYTHING FROM THIS LINE ON ! # # DO NOT CHANGE ANYTHING FROM THIS LINE ON ! #

$ScriptVersion = "V2.0.0" # 06.10.2024

#####################################################################################################
# Was the script started with Administrator priviliges?:
#####################################################################################################
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
	Start-Process powershell.exe "-WindowStyle Maximized -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
	Exit
}


#####################################################################################################
# Welcome text outout
#####################################################################################################
$ScriptSteps = 12
clear
Write-Host "
      ___           ___           ___           ___           ___           ___       ___           ___           ___     
     /\__\         /\  \         /\  \         /\  \         /\  \         /\__\     /\  \         /\  \         /\__\    
    /:/ _/_       /::\  \       /::\  \       /::\  \       /::\  \       /:/  /    /::\  \       /::\  \       /:/  /    
   /:/ /\__\     /:/\:\  \     /:/\:\  \     /:/\:\  \     /:/\:\  \     /:/  /    /:/\:\  \     /:/\:\  \     /:/__/     
  /:/ /:/ _/_   /:/  \:\  \   /::\~\:\  \   /:/  \:\__\   /:/  \:\  \   /:/  /    /:/  \:\  \   /:/  \:\  \   /::\__\____ 
 /:/_/:/ /\__\ /:/__/ \:\__\ /:/\:\ \:\__\ /:/__/ \:|__| /:/__/ \:\__\ /:/__/    /:/__/ \:\__\ /:/__/ \:\__\ /:/\:::::\__\
 \:\/:/ /:/  / \:\  \ /:/  / \/_|::\/:/  / \:\  \ /:/  / \:\  \  \/__/ \:\  \    \:\  \ /:/  / \:\  \  \/__/ \/_|:|~~|~   
  \::/_/:/  /   \:\  /:/  /     |:|::/  /   \:\  /:/  /   \:\  \        \:\  \    \:\  /:/  /   \:\  \          |:|  |    
   \:\/:/  /     \:\/:/  /      |:|\/__/     \:\/:/  /     \:\  \        \:\  \    \:\/:/  /     \:\  \         |:|  |    
    \::/  /       \::/  /       |:|  |        \::/__/       \:\__\        \:\__\    \::/  /       \:\__\        |:|  |    
     \/__/         \/__/         \|__|         ~~            \/__/         \/__/     \/__/         \/__/         \|__|    
"
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host "Automatic Arduino IDE and library setup for WordClock by AWSW -" $ScriptVersion
Write-Host " " 
Write-Host "Installing Arduino IDE, required libraries and download code for your WordClock fully automatically."
Write-Host " " 
Write-Host "! The installation is done hard! This meaans existing Arduino IDE or library installations get removed hard !"
Write-Host "! This ensures that the compile will work after this script was used but can cause the removal of old libraries !"
Write-Host "! Stop this script now if you want to save old libraries first !" 
Write-Host "! IMPORTANT: You are using this script at your own risk. The manual way to do this is well described too !" 
Write-Host " "
Write-Host " "
Write-Host "Everything will be downloaded and installed for you... You just need to wait some minutes until the process is finished... =)"
Write-Host " "
Write-Host " "
Write-Host "Just wait! Several times the Arduino IDE will open and close automatically. Do not use your keyboard or mouse during this time!"
Write-Host " "
Write-Host " "
Write-Host "Step numbers will be shown to show you some progress. The skript runs as fast as possible, depending on your machine speed. WAIT!"
Write-Host " "
Write-Host " "
Write-Host "Please press ENTER and JUST WAIT until the process is done..."
Write-Host " " 
pause
clear

Sleep 1
clear
Write-Host "Starting in 5... Please just wait... Do not use your keyboard or mouse during this time!"
Sleep 1
clear
Write-Host "Starting in 4... Please just wait... Do not use your keyboard or mouse during this time!"
Sleep 1
clear
Write-Host "Starting in 3... Please just wait... Do not use your keyboard or mouse during this time!"
Sleep 1
clear
Write-Host "Starting in 2... Please just wait... Do not use your keyboard or mouse during this time!"
Sleep 1
clear
Write-Host "Starting in 1... Please just wait... Do not use your keyboard or mouse during this time!"
Sleep 1
clear

#####################################################################################################
# Automatic cleanup of old Arduino IDE and code folders:
#####################################################################################################
#   All users: C:\Program Files\Arduino IDE
#   Only me: C:\Users\%USERNAME%\AppData\Local\Programs\Arduino IDE
#   C:\Users\{username}\AppData\Local\Arduino15
#   C:\Users\{username}\.arduinoIDE
#   C:\Users\{username}\AppData\Roaming\arduino-ide\

# Destination folders:
$DestinationFolder0 = "$env:USERPROFILE\Downloads\AWSW-WordClock-Code"
$DestinationFolder1 = "$env:USERPROFILE\Downloads\WordClock-TEMP"
$DestinationFolder2 = "$env:USERPROFILE\Documents\Arduino\libraries"
$DestinationFolder3 = "$env:USERPROFILE\.arduinoIDE"
$DestinationFolder4 = "$env:USERPROFILE\AppData\Local\Arduino15"
$DestinationFolder5 = "$env:USERPROFILE\AppData\Roaming\arduino-ide\"


clear
# Remove existing folders:
Write-Host " "
Write-Host "STEP 1 OF" $ScriptSteps "- Automatic cleanup of old Arduino IDE and code folders:"
Write-Host " "
Write-Host "Automatic delete of folder: $DestinationFolder0"
Remove-Item $DestinationFolder0 -Recurse -ErrorAction Ignore
Write-Host " "
Sleep 1
Write-Host "Automatic delete of folder: $DestinationFolder1"
Remove-Item $DestinationFolder1 -Recurse -ErrorAction Ignore
Write-Host " "
Sleep 1
Write-Host "Automatic delete of folder: $DestinationFolder2"
Remove-Item $DestinationFolder2 -Recurse -ErrorAction Ignore
Write-Host " "
Sleep 1
Write-Host "Automatic delete of folder: $DestinationFolder3"
Remove-Item $DestinationFolder3 -Recurse -ErrorAction Ignore
Write-Host " "
Sleep 1
Write-Host "Automatic delete of folder: $DestinationFolder4"
Remove-Item $DestinationFolder4 -Recurse -ErrorAction Ignore
Write-Host " "
Sleep 1
Write-Host "Automatic delete of folder: $DestinationFolder5"
Remove-Item $DestinationFolder5 -Recurse -ErrorAction Ignore
Write-Host " "
Sleep 1

clear

# Ensure the destination folder exists
If (-not (Test-Path $DestinationFolder0)) {
    New-Item -Path $DestinationFolder0 -ItemType Directory
} 
If (-not (Test-Path $DestinationFolder1)) {
    New-Item -Path $DestinationFolder1 -ItemType Directory
} 
If (-not (Test-Path $DestinationFolder2)) {
    New-Item -Path $DestinationFolder2 -ItemType Directory
} 
clear


#####################################################################################################
# Automatic uninstall of old Arduino IDE installations:
#####################################################################################################
Clear
Write-Host " "
Write-Host "STEP 2 OF" $ScriptSteps "- Automatic uninstall of old Arduino IDE installation start"
Write-Host " "
Write-Host "Please wait! This might take some time... Do not stop the script!"
Write-Host " "
Invoke-Command -ScriptBlock {Get-CimInstance -ClassName Win32_Product -Filter "name like 'Arduino IDE'" | Invoke-CimMethod -MethodName Uninstall}
Write-Host "Automatic uninstall of old Arduino IDE installation done"
Write-Host " "
Sleep 3


#####################################################################################################
# Automatic Arduino IDE .MSI download:
#####################################################################################################
clear
Write-Host " "
Write-Host "STEP 3 OF" $ScriptSteps "- Downloading file: Arduino IDE 2.3.3" 
Write-Host " "
Write-Host "Please wait! This might take some time... Do not stop the script!"
Write-Host " "
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -Uri "https://downloads.arduino.cc/arduino-ide/arduino-ide_2.3.3_Windows_64bit.msi" -OutFile "$DestinationFolder1\arduino-ide_Windows_64bit.msi"
Sleep 3


#####################################################################################################
# Automatic Arduino IDE setup installation:
#####################################################################################################
clear
Write-Host " "
Write-Host "STEP 4 OF" $ScriptSteps "- Automatic install of new Arduino IDE 2.3.3:"
Write-Host " "
Start-Process -FilePath "$DestinationFolder1\arduino-ide_Windows_64bit.msi" -ArgumentList "/qb /norestart" -Wait
Sleep 3


#####################################################################################################
# Start and quit Arduino IDE once:
#####################################################################################################
clear
Sleep 1
Write-Host "!!! Now Arduino IDE will be started automatically. Wait until it is fully loaded and closed automatically. "
Write-Host " "
Write-Host "! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything !"
Write-Host " "
Write-Host "STEP 5 OF" $ScriptSteps "- Starting in 5"
Write-Host " "
Sleep 1
clear
Write-Host "!!! Now Arduino IDE will be started automatically. Wait until it is fully loaded and closed automatically. "
Write-Host " "
Write-Host "! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything !"
Write-Host " "
Write-Host "STEP 5 OF" $ScriptSteps "- Starting in 4"
Write-Host " "
Sleep 1
clear
Write-Host "!!! Now Arduino IDE will be started automatically. Wait until it is fully loaded and closed automatically. "
Write-Host " "
Write-Host "! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything !"
Write-Host " "
Write-Host "STEP 5 OF" $ScriptSteps "- Starting in 3"
Write-Host " "
Sleep 1
clear
Write-Host "!!! Now Arduino IDE will be started automatically. Wait until it is fully loaded and closed automatically. "
Write-Host " "
Write-Host "! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything !"
Write-Host " "
Write-Host "STEP 5 OF" $ScriptSteps "- Starting in 2"
Write-Host " "
Sleep 1
clear
Write-Host "!!! Now Arduino IDE will be started automatically. Wait until it is fully loaded and closed automatically. "
Write-Host " "
Write-Host "! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything !"
Write-Host " "
Write-Host "STEP 5 OF" $ScriptSteps "- Starting in 1"
Write-Host " "
Sleep 1
clear

$programtorun = "$env:USERPROFILE\AppData\Local\Programs\arduino-ide\Arduino IDE.exe"
$programtorunpath = "$env:USERPROFILE\AppData\Local\Programs\arduino-ide"

for ($i=1; $i -le 1; $i++)
{
    $proc = Start-Process -filePath $programtorun -workingdirectory $programtorunpath -PassThru -WindowStyle Normal
    # keep track of timeout event
    $timeouted = $null # reset any previously set timeout
    # wait up to x seconds for normal termination
    $proc | Wait-Process -Timeout 60 -ErrorAction SilentlyContinue -ErrorVariable timeouted
    if ($timeouted)
    {
        # terminate the process
        $proc | kill
        # update internal error counter
    }
    elseif ($proc.ExitCode -ne 0)
    {
        # update internal error counter
    }
}


#####################################################################################################
# Modify the board manager list to be able to add ESP32 boards:
#####################################################################################################
clear
Write-Host " "
Write-Host "STEP 6 OF" $ScriptSteps "- Setting the board managers for ESP32 boards in the Arduino IDE configuration"
Write-Host " "
Sleep 3
$fileName = "$env:USERPROFILE\.arduinoIDE\arduino-cli.yaml"
(Get-Content -Path $fileName).Replace('  additional_urls: []','  additional_urls:') | Set-Content $fileName
(Get-Content $fileName) | 
    Foreach-Object {
        $_ # send the current line to output
        if ($_ -match "  additional_urls:") 
        {
  #Add lines after the selected pattern:
  "     - https://dl.espressif.com/dl/package_esp32_index.json"
        }
    } | Set-Content $fileName


#####################################################################################################
# Turn off Arduino IDE automatic updates to prevent version incompatibilities:
#####################################################################################################
clear
Write-Host " "
Write-Host "STEP 7 OF" $ScriptSteps "- Turn off Arduino IDE automatic updates to prevent version incompatibilities"
Write-Host " "
Sleep 3
$fileName = "$env:USERPROFILE\.arduinoIDE\settings.json"

$multilinetext = @"
{
  "window.titleBarStyle": "native",
  "workbench.colorTheme": "arduino-theme-dark",
  "editor.fontSize": 14,
  "files.autoSave": "afterDelay",
  "editor.quickSuggestions": {
    "other": true,
    "comments": false,
    "strings": false
  },
  "arduino.window.autoScale": true,
  "window.zoomLevel": 0,
  "arduino.compile.verbose": false,
  "arduino.compile.warnings": "None",
  "arduino.upload.verbose": false,
  "arduino.upload.verify": false,
  "arduino.sketchbook.showAllFiles": false,
  "arduino.checkForUpdates": false
}
"@

$multiLineText | Out-File -FilePath $fileName



#####################################################################################################
# Start and quit Arduino IDE once to install the board managers for the ESP32:
#####################################################################################################
clear
Sleep 1
Write-Host "!!! Now Arduino IDE will be started automatically. Wait until it is fully loaded and closed automatically. "
Write-Host " "
Write-Host "! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything !"
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Starting in 5"
Write-Host " "
Sleep 1
clear
Write-Host "!!! Now Arduino IDE will be started automatically. Wait until it is fully loaded and closed automatically. "
Write-Host " "
Write-Host "! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything !"
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Starting in 4"
Write-Host " "
Sleep 1
clear
Write-Host "!!! Now Arduino IDE will be started automatically. Wait until it is fully loaded and closed automatically. "
Write-Host " "
Write-Host "! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything !"
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Starting in 3"
Write-Host " "
Sleep 1
clear
Write-Host "!!! Now Arduino IDE will be started automatically. Wait until it is fully loaded and closed automatically. "
Write-Host " "
Write-Host "! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything !"
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Starting in 2"
Write-Host " "
Sleep 1
clear
Write-Host "!!! Now Arduino IDE will be started automatically. Wait until it is fully loaded and closed automatically. "
Write-Host " "
Write-Host "! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything ! Do not do anything !"
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Starting in 1"
Write-Host " "
Sleep 1
clear
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Starting Arudino IDE"
Write-Host " "
Sleep 3
Add-Type -AssemblyName System.Windows.Forms
$programtorun = "$env:USERPROFILE\AppData\Local\Programs\arduino-ide\Arduino IDE.exe"
$programtorunpath = "$env:USERPROFILE\AppData\Local\Programs\arduino-ide"
$proc = Start-Process -filePath $programtorun -workingdirectory $programtorunpath -PassThru -WindowStyle Normal
# Write-Host $proc.Id
clear
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Waiting for Arduino IDE startup"
Sleep 1
while (-not $wshell.AppActivate($proc.Id)) {
  Start-Sleep -MilliSeconds 200
}
clear
$wshell = New-Object -ComObject wscript.shell;
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Starting Arudino IDE"
Write-Host " "
Sleep 5
Add-Type -AssemblyName System.Windows.Forms
$programtorun = "$env:USERPROFILE\AppData\Local\Programs\arduino-ide\Arduino IDE.exe"
$programtorunpath = "$env:USERPROFILE\AppData\Local\Programs\arduino-ide"
$proc = Start-Process -filePath $programtorun -workingdirectory $programtorunpath -PassThru -WindowStyle Normal
clear
Sleep 5
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Activating Arduino IDE window and adding ESP32 board managers - this might take some minutes - please wait !!!"
Write-Host " "
Sleep 10
$wshell.AppActivate($proc.Id)
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys("^+{P}")
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys("boardmanager")
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys({~})
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys("{BACKSPACE 25}")
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys("esp32 espressif")
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys("{TAB 4}")
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys("{2}")
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys("{TAB}")
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys({~})
Sleep 3
$wshell.AppActivate($proc.Id)
$wshell.SendKeys("^+{P}")
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys("boardmanager")
Sleep 1
$wshell.AppActivate($proc.Id)
$wshell.SendKeys({~})
clear
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Activating Arduino IDE window and adding ESP32 board managers - this might take some minutes - please wait !!!"
Write-Host " "
Write-Host "... "
Write-Host " "
Sleep 60
clear
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Activating Arduino IDE window and adding ESP32 board managers - this might take some minutes - please wait !!!"
Write-Host " "
Write-Host "."
Write-Host " "
Sleep 30
clear
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Activating Arduino IDE window and adding ESP32 board managers - this might take some minutes - please wait !!!"
Write-Host " "
Write-Host ".."
Write-Host " "
Sleep 30
clear
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Activating Arduino IDE window and adding ESP32 board managers - this might take some minutes - please wait !!!"
Write-Host " "
Write-Host "..."
Write-Host " "
Sleep 30
clear
Write-Host " "
Write-Host "STEP 8 OF" $ScriptSteps "- Closing Arduino IDE window now automatically - please wait !!!"
Write-Host " "
Write-Host " "
Write-Host " "
Sleep 10
$wshell.AppActivate($proc.Id)
$wshell.SendKeys("%{F4}") # terminate the process
clear

#####################################################################################################
# Get the for WordClock required libraries:
#####################################################################################################
clear
Sleep 3
clear
Write-Host " "
Write-Host "STEP 9 OF" $ScriptSteps "- Downloading required library files and the WordClock code:"
Write-Host " "
Sleep 3
Write-Host "Downloading file: Adafruit_NeoPixel.zip" 
Invoke-WebRequest -Uri "https://codeload.github.com/adafruit/Adafruit_NeoPixel/zip/refs/heads/master" -OutFile "$DestinationFolder1\Adafruit_NeoPixel.zip"
Write-Host "Downloading file: AsyncTCP.zip"
Invoke-WebRequest -Uri "https://codeload.github.com/me-no-dev/AsyncTCP/zip/refs/heads/master" -OutFile "$DestinationFolder1\AsyncTCP.zip"
Write-Host "Downloading file: ESPAsyncWebServer.zip"
Invoke-WebRequest -Uri "https://codeload.github.com/me-no-dev/ESPAsyncWebServer/zip/refs/heads/master" -OutFile "$DestinationFolder1\ESPAsyncWebServer.zip" 
Write-Host "Downloading file: ESPUI.zip"
Invoke-WebRequest -Uri "https://github.com/s00500/ESPUI/archive/refs/tags/2.2.3.zip" -OutFile "$DestinationFolder1\ESPUI.zip" 
Write-Host "Downloading file: ArduinoJson.zip"
Invoke-WebRequest -Uri "https://codeload.github.com/bblanchon/ArduinoJson/zip/refs/heads/7.x" -OutFile "$DestinationFolder1\ArduinoJson.zip"
Write-Host "Downloading file: LITTLEFS.zip"
Invoke-WebRequest -Uri "https://codeload.github.com/lorol/LITTLEFS/zip/refs/heads/master" -OutFile "$DestinationFolder1\LITTLEFS.zip" 
Write-Host "Downloading file: Telegram-Bot.zip"
Invoke-WebRequest -Uri "https://codeload.github.com/witnessmenow/Universal-Arduino-Telegram-Bot/zip/refs/heads/master" -OutFile "$DestinationFolder1\Telegram-Bot.zip"
Write-Host "Downloading file: ESP32Time.zip"
Invoke-WebRequest -Uri "https://codeload.github.com/fbiego/ESP32Time/zip/refs/heads/main" -OutFile "$DestinationFolder1\ESP32Time.zip"
Write-Host "Downloading code for AWSW-WordClock models now:"
Write-Host "Downloading file: Code for AWSW WordClock 16x8 LED Matrix 2023"
Write-Host "Downloading file: Code for AWSW WordClock 16x8 LED Matrix 2024"
Invoke-WebRequest -Uri "https://codeload.github.com/AWSW-de/WordClock-16x16-LED-matrix-2023/zip/refs/heads/main" -OutFile "$env:USERPROFILE\Downloads\AWSW-WordClock-16x8-LED-matrix.zip"
Write-Host "Downloading file: Code for AWSW WordClock 16x16 LED Matrix (Smart variant with Telegram)"
Invoke-WebRequest -Uri "https://codeload.github.com/AWSW-de/WordClock-16x16-LED-matrix/zip/refs/heads/main" -OutFile "$env:USERPROFILE\Downloads\AWSW-WordClock-16x16-LED-matrix.zip"
Write-Host " "
Sleep 3

Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip
{
    param([string]$zipfile, [string]$outpath)

    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

clear
Write-Host " "
Write-Host "STEP 10 OF" $ScriptSteps "- Unzip library files to Arduino Library folder:"
Write-Host " "
Sleep 3
Unzip "$DestinationFolder1\Adafruit_NeoPixel.zip" $DestinationFolder2
Unzip "$DestinationFolder1\AsyncTCP.zip" $DestinationFolder2
Unzip "$DestinationFolder1\ESPAsyncWebServer.zip" $DestinationFolder2
Unzip "$DestinationFolder1\ESPUI.zip" $DestinationFolder2
Unzip "$DestinationFolder1\ArduinoJson.zip" $DestinationFolder2
Unzip "$DestinationFolder1\LITTLEFS.zip" $DestinationFolder2
Unzip "$DestinationFolder1\Telegram-Bot.zip" $DestinationFolder2
Unzip "$DestinationFolder1\ESP32Time.zip" $DestinationFolder2
Unzip "$env:USERPROFILE\Downloads\AWSW-WordClock-16x8-LED-matrix.zip" "$DestinationFolder0"
Unzip "$env:USERPROFILE\Downloads\AWSW-WordClock-16x16-LED-matrix.zip" "$DestinationFolder0"
clear


#####################################################################################################
# Automatic download and install ESP32 board driver:
#####################################################################################################
clear
Write-Host " "
Write-Host "STEP 11 OF" $ScriptSteps "- Automatic download and install ESP32 board driver..."
Write-Host " "
Sleep 3
Invoke-WebRequest -Uri "https://www.silabs.com/documents/public/software/CP210x_Universal_Windows_Driver.zip" -OutFile "$DestinationFolder1\CP210x_Universal_Windows_Driver.zip"
New-Item -Path "$DestinationFolder1\CP210x-Driver" -ItemType Directory
Unzip "$DestinationFolder1\CP210x_Universal_Windows_Driver.zip" "$DestinationFolder1\CP210x-Driver"
Sleep 1
Get-ChildItem "$DestinationFolder1\CP210x-Driver" -Recurse -Filter "*inf" | ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install }
Sleep 1


#####################################################################################################
# Show result:
#####################################################################################################
clear
Write-Host " "
Write-Host "STEP 12 OF" $ScriptSteps "- Cleaning up the temporary library download folder:"
Write-Host " "
Sleep 3
Remove-Item $DestinationFolder1 -Recurse
Remove-Item "$env:USERPROFILE\Downloads\AWSW-WordClock-16x8-LED-matrix.zip"
Remove-Item "$env:USERPROFILE\Downloads\AWSW-WordClock-16x16-LED-matrix.zip"
Sleep 1
clear
Write-Host "
      ___           ___           ___           ___           ___           ___       ___           ___           ___     
     /\__\         /\  \         /\  \         /\  \         /\  \         /\__\     /\  \         /\  \         /\__\    
    /:/ _/_       /::\  \       /::\  \       /::\  \       /::\  \       /:/  /    /::\  \       /::\  \       /:/  /    
   /:/ /\__\     /:/\:\  \     /:/\:\  \     /:/\:\  \     /:/\:\  \     /:/  /    /:/\:\  \     /:/\:\  \     /:/__/     
  /:/ /:/ _/_   /:/  \:\  \   /::\~\:\  \   /:/  \:\__\   /:/  \:\  \   /:/  /    /:/  \:\  \   /:/  \:\  \   /::\__\____ 
 /:/_/:/ /\__\ /:/__/ \:\__\ /:/\:\ \:\__\ /:/__/ \:|__| /:/__/ \:\__\ /:/__/    /:/__/ \:\__\ /:/__/ \:\__\ /:/\:::::\__\
 \:\/:/ /:/  / \:\  \ /:/  / \/_|::\/:/  / \:\  \ /:/  / \:\  \  \/__/ \:\  \    \:\  \ /:/  / \:\  \  \/__/ \/_|:|~~|~   
  \::/_/:/  /   \:\  /:/  /     |:|::/  /   \:\  /:/  /   \:\  \        \:\  \    \:\  /:/  /   \:\  \          |:|  |    
   \:\/:/  /     \:\/:/  /      |:|\/__/     \:\/:/  /     \:\  \        \:\  \    \:\/:/  /     \:\  \         |:|  |    
    \::/  /       \::/  /       |:|  |        \::/__/       \:\__\        \:\__\    \::/  /       \:\__\        |:|  |    
     \/__/         \/__/         \|__|         ~~            \/__/         \/__/     \/__/         \/__/         \|__|    
"
Write-Host " "
Write-Host "Downloading and installing of Arduino IDE fully done. Required library files and the WordClock code were downloaded and unziped."
Write-Host " "
Write-Host "The code for the WordClock is stored in: '$DestinationFolder0'"
Write-Host "The folder will be opened after you press ENTER after your read and confirmed this text."
Write-Host " "
Write-Host " "
Write-Host "Steps you need to do manually in the Arduino IDE now:"
Write-Host "- Open the 'Arduino IDE'."
Write-Host "- Select the 'WEMOS D1 MINI ESP32' (2023 models) or 'NodeMCU-32S' (2024 models) from the menu 'Select other board and port' and select the port the ESP32 is connected to."
Write-Host "- Open, compile and upload the WordClock code of your variante stored in: '$DestinationFolder0'"
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host "Enjoy your WordClock =)"
Write-Host " "
Write-Host "Kind regards AWSW"
Write-Host " "
Write-Host " "
pause
Start-Process -FilePath "$DestinationFolder0"
