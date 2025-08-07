$app = winget list "<APPID>" -e --accept-source-agreements 2>&1

# Use regex to extract installed version
$regex = [regex]::Match($app, "<APPID>\s+([\d\.]+)\s+[\d\.]+\s+winget")

if ($regex.Success) {
    $installedVersion = [Version]$regex.Groups[1].Value
    $requiredVersion = [Version]"<APPVERSION>"

    if ($installedVersion -ge $requiredVersion) {
        Write-Output "Detected version: $installedVersion (meets or exceeds required: $requiredVersion)"
        exit 0
    } else {
        Write-Output "Installed version $installedVersion is older than required $requiredVersion"
        exit 1
    }
} else {
    Write-Output "Application not found or regex mismatch"
    exit 1
}
