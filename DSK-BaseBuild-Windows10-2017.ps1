## Created by: Nicholas Glantz
## Function: Cleans up Windows 10 start menu tiles
##           Removes any version of Office currently installed on the machine.
##           Installs the McAfee Agent
##           Installs DSK Base apps
## Date: December 6th, 2017
## Location: #

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force
Add-Type -AssemblyName PresentationFramework

## Variables
$startlayout = 'startlayout.bin'
$rootdirectory = 'C:\'
$internetexplorer = 'C:\Users\Administrator\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Accessories\Internet Explorer.lnk'
$startlayoutdirectory = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories'
$sysprepscript = 'RunPriorToSysprep.ps1'
$sysprepdirectory = 'C:\Windows\System32\Sysprep'

## Copy and Paste
copy -Path $internetexplorer -Destination $startlayoutdirectory
copy -Path $startlayout -Destination $rootdirectory
copy -Path $sysprepscript -Destination $sysprepdirectory

## Run Windows10-Initial-Setup.ps1
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force
& 'Windows10-Initial-Setup.ps1'

##removes any version of Office currently installed on the pc
Start-Process 'O15CTRRemove.diagcab' -Wait

##Installs base apps 
Start-Process 'DSK-BUILD_AUTUM_2016.BAT' -Wait

##Installs McAfee Agent
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force
& 'Install.ps1'

##If base app installation wasn't succesful prompt user for retry.
if(!( (Test-Path 'C:\Program Files\7-Zip') -and (Test-Path 'C:\Program Files (x86)\Adobe') -and (Test-Path 'C:\Program Files\Microsoft Silverlight') -and (Test-Path 'C:\Program Files (x86)\K-Lite Codec Pack' ) )){
    $erroraction = [System.Windows.MessageBox]::Show('There was an error installing base app package. Would you like to retry?','Error!', 'YesNo','Error')
    switch ($erroraction){
    'Yes' {
            rmdir "C:\LOGS\*" -ErrorAction SilentlyContinue
            Start-Process '\\itres\deploy$\BUILD\DSK-BUILD_AUTUM_2016.BAT' -Wait
           }

    'No' {
        
         }
    }
}
 

[System.Windows.MessageBox]::Show('The build script succesfully executed!','Success!','Ok','Information')