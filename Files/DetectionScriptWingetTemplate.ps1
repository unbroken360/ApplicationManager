$app = winget list "<APPID>" -e --accept-source-agreements 2>&1

# Use regex to extract installed version
$regex = [regex]::Match($app, "<APPID>\s+([\d\.]+)\s+[\d\.]+\s+winget")

if ($regex.Success) {
    $installedVersion = $regex.Groups[1].Value

    if ($installedVersion -eq "<APPVERSION>") {
        Write-Output "Detected correct version: $installedVersion"
        exit 0
    } else {
        Write-Output "Wrong version installed: $installedVersion"
        exit 1
    }
} else {
    Write-Output "Application not found or regex mismatch"
    exit 1
}
