# Prompt for starting search-id
$startId = Read-Host "Enter starting search-id"
$searchId = [int]$startId

# Output CSV file
$outputCsv = "Workflows.csv"

# Define CSV header if file doesn't exist
if (-not (Test-Path $outputCsv)) {
    "address,search-id,name,search-string,category-id,frequently-used,mail-address,smb-directory,smb-username,smb-password" | Out-File -FilePath $outputCsv -Encoding UTF8
}

# Process each WF_*.txt file
Get-ChildItem -Filter "WF_*.txt" | ForEach-Object {
    $content = Get-Content $_.FullName -Raw

    $name = [regex]::Match($content, "Distribution:\s+(.*)").Groups[1].Value.Trim()
    $ipAddressPort = [regex]::Match($content, "IP Address:Port\s+(.*)").Groups[1].Value.Trim()
    $ipAddress = $ipAddressPort -replace ":\d+$", ""
    $share = [regex]::Match($content, "Share:\s+(.*)").Groups[1].Value.Trim()
    $docPath = [regex]::Match($content, "Document Path:\s+(.*)").Groups[1].Value.Trim() -replace "[/\\]", "\"
    $loginName = [regex]::Match($content, "Login Name:\s+(.*)").Groups[1].Value.Trim()

    $uncPath = "\\$ipAddress\$share\$docPath\"

    $searchString = if ($name.Length -gt 10) { $name.Substring(0,10) } else { $name }

    $csvLine = '"data",{0},"{1}","{2}",2,True,,"{3}","{4}",""' -f $searchId, $name, $searchString, $uncPath, $loginName
    Add-Content -Path $outputCsv -Value $csvLine

    $searchId++
}
