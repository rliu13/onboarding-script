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
)

# Remove bloatware
foreach ($app in $bloatware) {
    Get-AppxPackage -Name $app | Remove-AppxPackage
    Get-AppxProvisionedPackage -Online | where {$_.DisplayName -like $app} | Remove-AppxProvisionedPackage -Online
}

# Install Microsoft Office if a CDN is provided
if($env:OfficeCDN -ne $null) {
    try {
        $webClient = New-Object System.Net.WebClient

        # Use the user's temp directory
        $tempFolder = [System.IO.Path]::GetTempPath()
        # $destination = [System.IO.Path]::Combine($tempFolder, "Office.img")
	$destination = Join-Path -Path $tempFolder -ChildPath "MicrosoftOffice.img"
        $webClient.DownloadFile($env:OfficeCDN, $destination)

        # Mount the image
        $mountResult = Mount-DiskImage -ImagePath $destination -PassThru

        $driveLetter = ($mountResult | Get-Volume).DriveLetter

        # Start Setup.exe
        Start-Process -FilePath "Setup.exe" -WorkingDirectory "$($driveLetter):\" -Wait
    
        # Dismount image
        Dismount-DiskImage -ImagePath $destination

        # Delete file
        Remove-Item -Path $destination
    }
    catch [System.Net.WebException],[System.IO.IOException] {
    Write-Error "Unable to download Microsoft Office from $($env:OfficeCDN)"
    }
    catch {
    Write-Error "An unresolved error occurred"
    }
}

