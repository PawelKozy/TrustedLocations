$userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false }
foreach($userProfile in $userProfiles)
{
    $user = $userProfile.LocalPath
    $sid = $userProfile.SID
    $officeApps = @("Word","Excel")
    foreach ($app in $officeApps)
    {
        $key = "HKEY_USERS\$sid\Software\Microsoft\Office\16.0\$app\Security\Trusted Locations"
        try {
            if (Test-Path $key) {
                $locationKeys = (Get-ChildItem -Path registry::$key -Recurse) | Where-Object { $_.Name-match "Location*" }
                if ($locationKeys) {
                    $TrustedLocations = @()
                    foreach($locationKey in $locationKeys){
                        $locationValue = (Get-ItemProperty -Path $locationKey.PSPath).Path
                        $TrustedLocations += $locationValue
                    }
                    if($TrustedLocations.Count -gt 0){
                        Write-Host "`nUser: $user `nSID: $sid" -ForegroundColor Cyan
                        foreach($location in $TrustedLocations){
                            Write-Host "`nTrusted Location: $location" -ForegroundColor Green
                        }
                    }else{
                        Write-Host "No trusted locations found for user $user with SID $sid"
                    }
                } else {
                    Write-Host "No Trusted Locations found for user $user with SID $sid"
                }
            }
        } catch {
            Write-Host "An error occurred while trying to access the key $key. The error message is: $_" -ForegroundColor Red
        }
    }
}
