function Get-WingetCmd {

    #WinGet Path (if User/Admin context)
    $UserWingetPath = Get-Command winget.exe -ErrorAction SilentlyContinue
    #WinGet Path (if system context)
    $SystemWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe"

    #Get Winget Location in User/Admin context
    if ($UserWingetPath) {
        return $UserWingetPath.Source
    }
    #Get Winget Location in System context
    elseif ($SystemWingetPath) {
        #If multiple version, pick last one
        return $SystemWingetPath[-1].Path
    }
    else { 
        Write-Host "WinGet is not installed, mandatory to run WinGet integration"
        exit 1
    }

}
function Normalize-Version($ver) {
    # Wenn keine Punkte fehlen, auffüllen bis 4 Stellen
    $parts = $ver.Split('.')
    while ($parts.Count -lt 4) {
        $parts += '0'
    }
    return ($parts -join '.')
}

$Winget = Get-WingetCmd
$app = &($Winget) list "<APPID>" -e --accept-source-agreements 2>&1

# Use regex to extract installed version
$regex = [regex]::Match($app, "<APPID>\s+([\d\.]+)\s+(?:[\d\.]+\s+)?winget")

if ($regex.Success) {
    $installedVersion = [Version](Normalize-Version $regex.Groups[1].Value)
    $requiredVersion = [Version](Normalize-Version "<APPVERSION>")

    if ($installedVersion -ge $requiredVersion) {
        Write-Output "Detected version: $installedVersion (meets or exceeds required: $requiredVersion)"
        exit 0
    } else {
        # Installed version is older than required
        exit 1
    }
} else {
    # Application not found or regex mismatch
    exit 1
}
