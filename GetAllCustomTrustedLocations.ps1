$userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false }
foreach($userProfile in $userProfiles)
{
    $user = $userProfile.LocalPath
    $officeApps = @("Word", "Excel", "PowerPoint", "Outlook")
    foreach ($app in $officeApps) {
        $key = "HKCU:\Software\Microsoft\Office\16.0\$app\Security\Trusted Locations"
        if(Test-Path $key){
            $locationKeys = Get-ChildItem -Path $key | Where-Object { $_.PSChildName -match "Location*" }
            if ($locationKeys) {
                $TrustedLocations = @()
                foreach($locationKey in $locationKeys){
                    $TrustedLocations += (Get-ItemProperty -Path $locationKey.PSPath).Path
                }
                if($TrustedLocations.Count -gt 0){
                    foreach($location in $TrustedLocations){
                        Write-Host "`nUser: $user `nApplication: $app `nTrusted Location: $location" -ForegroundColor Green
                    }
                }else{
                    Write-Host "No trusted locations found for $app for user $user"
                }
            } else {
                Write-Host "No Trusted Locations found for $app for user $user"
            }
        } else {
            Write-Host "No Trusted Locations key found for $app for user $user"
        }
    }
}
