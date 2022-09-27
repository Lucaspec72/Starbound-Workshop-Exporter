#Starbound Workshop Exporter v1
#Made by Lucaspec72

#DEFAULT SETTINGS (will be used if user cancels folderbrowsedialog)
$workshopFolder = "C:\Program Files (x86)\Steam\steamapps\workshop\content\211820"
$outputFolder = "C:\Starbound Workshop Exporter"
Clear-Host
#Checks if steam is available
$steamAccess = 0
Write-Host "Trying to reach steamcommunity.com..."
#Increasing Count value or TTL might help if connection is too slow.
if(Test-Connection "steamcommunity.com" -Count 4 -TimeToLive 135){
    $steamAccess = 1
    Write-Host "Successfully connected to Steam Workshop !" -BackgroundColor Green -ForegroundColor Black
} else {Write-Host "Could not connect to Steam Workshop, running in offline mode." -BackgroundColor Yellow -ForegroundColor Black}

#FolderBrowserDialog object creation
Add-Type -AssemblyName System.Windows.Forms
$browser = New-Object System.Windows.Forms.FolderBrowserDialog
$browser.RootFolder = "MyComputer"
$browser.Description = "Please select the location of your starbound workshop folder. (default : C:\Program Files (x86)\Steam\steamapps\workshop\content\211820)"
Write-Host "Folder Browser window opened, check your background windows..." -ForegroundColor Black -BackgroundColor DarkCyan
Write-Host "Waiting for workshop folder select window..." -ForegroundColor Black -BackgroundColor DarkCyan
$null = $browser.ShowDialog()
$path = $browser.SelectedPath
Try {
if ([string]::IsNullOrEmpty($path) -and (Test-Path $workshopFolder)) {
    Write-Host "No path set but found folder on default install location, proceeding with default parameters." -ForegroundColor yellow
} elseif (Test-Path $path){
    $workshopFolder = $path
    Write-Host "Workshop folder set to '$path'" -ForegroundColor green
}
} Catch {
    Write-Host "Workshop folder location unknown, aborting execution." -ForegroundColor red
    exit
}
Clear-Host
$mods = Get-ChildItem -Path $workshopFolder -Directory | Select-Object Select -expa Name
foreach ($mod in $mods){
    $friendlyModName = $mod
    if($steamAccess){
        #The following code gets the human-friendly name of the mod from the modID
        Try{
        $URI = "https://steamcommunity.com/sharedfiles/filedetails/?id=$mod"
        $webpage = Invoke-WebRequest -Uri $URI -UseBasicParsing
        $html = New-Object -ComObject "HTMLFile"
        $src = [System.Text.Encoding]::Unicode.GetBytes($webpage.Content)
        $html.write($src)
        $workshopItemTitle = $html.body.getElementsByClassName('workshopItemTitle')[0]
        $friendlyModName = $workshopItemTitle.innerText
        $origFriendlyModName = $friendlyModName
        #Remove special Characters
        [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {$friendlyModName = $friendlyModName.replace($_,' ')}
        } catch {
            Write-Host "`nThere was a error retreiving mod info from steam for mod $mod" -ForegroundColor Red
        }
    }
    Try {
    Copy-Item "$workshopFolder\$mod\*.pak" "$outputFolder\$friendlyModName.pak" -ErrorAction Stop
    Write-Host "Mod Exported : $origFriendlyModName"
    } Catch {
    Write-Host "`nERROR, the following folder has non-standard folder structure" -ForegroundColor red
    Write-Host "$friendlyModName" -ForegroundColor yellow
    Write-Host "attempting copy of folder to output folder..." -ForegroundColor gray
    Try {
        Copy-Item "$workshopFolder\$mod" "$outputFolder\$mod" -Recurse -ErrorAction Stop
        Write-Host "folder has been copied to output folder." -ForegroundColor green
    } Catch {
        Write-Host "folder could not be copied, investigate manually." -ForegroundColor red
    }
    Write-Host " `n`n"
    }
}
Write-Host "Done !" -ForegroundColor Black -BackgroundColor Green
explorer.exe $outputFolder