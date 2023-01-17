$userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false }
foreach($userProfile in $userProfiles)
{
    $user = $userProfile.LocalPath
    $officeApps = @("Word", "Excel", "PowerPoint", "Outlook")
    $defaultLocations = @("C:\Users\$user\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Office",
                          "C:\Program Files (x86)\Microsoft Office",
                          "C:\Program Files\Microsoft Office")
    foreach ($app in $officeApps) {
        $key = "HKCU:\Software\Microsoft\Office\16.0\$app\Security\Trusted Locations"
        if(Test-Path $key){
            $locationKeys = Get-ChildItem -Path $key | Where-Object { $_.PSChildName -match "Location*" }
            if ($locationKeys) {
                $i = 1
                $TrustedLocations = @()
                foreach($locationKey in $locationKeys){
                    $TrustedLocations += (Get-ItemProperty -Path $locationKey.PSPath).Path
                }
                $CustomLocations = @()
                foreach($location in $TrustedLocations){
                    $match = $false
                    foreach($defaultLocation in $defaultLocations)
                    {
                        if ($location -like "$defaultLocation*")
                        {
                            $match = $true
                            break
                        }
                    }
                    if (!$match)
                    {
                        $CustomLocations += $location
                    }
                }
                if($CustomLocations.Count -gt 0){
                    foreach($location in $CustomLocations){
                        Write-Host "`nUser: $user `nApplication: $app `nCustom Trusted Location: $location" -ForegroundColor Green
                    }
                }else{
                    Write-Host "No custom trusted locations found for $app for user $user"
                }
            } else {
                Write-Host "No Trusted Locations found for $app for user $user"
            }
        } else {
            Write-Host "No Trusted Locations key found for $app for user $user"
        }
    }
}
