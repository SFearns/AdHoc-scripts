# This script will keep "pressing" the Scroll Lock key every 10 seconds
# The key I will use is the Scroll Lock key

# Change the Window title
$host.UI.RawUI.WindowTitle = "Keep Awake v2024.08.30"

Write-Host "`nStopping the screen saver, screen lock, etc"
Write-Host "`nPress Ctrl-C to cancel"

$temp = New-Object -ComObject "WScript.Shell"
While ($True) {
    # Sent twice to toggle the setting
    $temp.SendKeys("{ScrollLock}")
    $temp.SendKeys("{ScrollLock}")

    # Pause for 10 seconds
    Start-Sleep -Seconds 10

    # Constantly repeat until the User ends the script 
}