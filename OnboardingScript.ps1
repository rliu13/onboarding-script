# Get CDN url
$cdn = $env:OfficeCDN

# Set local support account's password to never expire
Set-LocalUser -Name "Support" -PasswordNeverExpires:$true

# Set time zone to Eastern Standard Time 
if((Get-TimeZone).Id -ne ("Eastern Standard Time") ) {
    Set-TimeZone -Name "Eastern Standard Time"
}

# Set computer to never sleep when plugged in
powercfg -change -standby-timeout-ac 0 

# Add bloatware here
$bloatware = @(
    'Microsoft.Xbox.TCUI'
    'Microsoft.XboxGamingOverlay'
    'Microsoft.XboxGameOverlay'
    'Microsoft.XboxIdentityProvider'
    'Microsoft.XboxSpeechToTextOverlay'
    'Microsoft.GamingApp'
    'Microsoft.ZuneMusic'
    'Microsoft.ZuneVideo'
    'Microsoft.BingNews'
    'Microsoft.BingWeather'
    'AD2F1837.HPDesktopSupportUtilities'
    'AppUp.IntelGraphicsExperience'
    'AppUp.IntelManagementandSecurityStatus'
    'RealtekSemiconductorCorp.HPAudioControl'
)

# Remove bloatware
foreach ($app in $bloatware) {
    Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -AllUsers
    Get-AppxProvisionedPackage -Online | where {$_.DisplayName -like $app} | Remove-AppxProvisionedPackage -Online
}

# Install office if a CDN link is provided
if($cdn -ne $null) {
    try {
        $webClient = New-Object System.Net.WebClient

        # Use temp folder
        $tempFolder = [System.IO.Path]::GetTempPath()
	    $destination = Join-Path -Path $tempFolder -ChildPath "Office.img"
        $webClient.DownloadFile($cdn, $destination)

        # Mount the image
        $mountResult = Mount-DiskImage -ImagePath $destination -PassThru

        $driveLetter = ($mountResult | Get-Volume).DriveLetter

        # Start Setup.exe
        Start-Process -FilePath "Setup.exe" -WorkingDirectory "$($driveLetter):\" -Wait
    
        # Dismount image
        Dismount-DiskImage -ImagePath $destination

        # Clean up temp folder
        Remove-Item -Path $destination
    }
    catch [System.Net.WebException],[System.IO.IOException] {
    Write-Error "Unable to download Microsoft Office from $($cdn)"
    }
    catch {
    Write-Error "An unresolved error occurred"
    }
}
