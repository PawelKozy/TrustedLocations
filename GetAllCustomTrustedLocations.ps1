$officeApps = @("Word", "Excel", "PowerPoint", "Outlook")
$user = (Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false }).LocalPath
$defaultLocations = @("C:\Users\$user\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Office",
                      "C:\Program Files (x86)\Microsoft Office",
                      "C:\Program Files\Microsoft Office")
                      
foreach ($app in $officeApps) {
    $key = "HKLM:\Software\Microsoft\Office\16.0\$app\Security\Trusted Locations"
    if(Test-Path $key){
        $locationKeys = Get-ChildItem -Path $key | Where-Object { $_.PSChildName -match "Location*" }
        Write-Host "`nTrusted Locations for ${app}:" -ForegroundColor Cyan
        if ($locationKeys) {
            $i = 1
            $TrustedLocations = @()
            foreach($locationKey in $locationKeys){
                $TrustedLocations += (Get-ItemProperty -Path $locationKey.PSPath).Path
            }
            $CustomLocations = Compare-Object -ReferenceObject $defaultLocations -DifferenceObject $TrustedLocations | Select-Object -ExpandProperty InputObject
            if($CustomLocations.Count -gt 0){
                foreach($location in $CustomLocations){
                    Write-Host "`n$location" -ForegroundColor Green
                }
            }else{
                Write-Host "No custom trusted locations found for $app"
            }
        } else {
            Write-Host "No Trusted Locations found for $app"
        }
    } else {
        Write-Host "No Trusted Locations key found for $app"
    }
}
