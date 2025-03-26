Add-Type -AssemblyName System.Windows.Forms

function Get-InputFilePath {
    Write-Host "Please make sure that your CSV has a header/starts with ""ComputerName"" *CASE SENSITIVE*"
    Write-Host "A prompt will open within 10 seconds, please enter your input CSV file"
    Start-Sleep -Seconds 5
    $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $FileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $FileDialog.Title = "Select Input CSV File"
    if ($FileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $FileDialog.FileName
    } else {
        Write-Host "No input file selected. Exiting."
        exit
    }
}

$InputCSV = Get-InputFilePath

Import-Module ActiveDirectory

$Computers = Import-Csv -Path $InputCSV

$TotalCount = 0
$InDomainCount = 0
$NotInDomainCount = 0
$InDomainComputers = @() # List to store computers in AD

foreach ($Computer in $Computers) {
    $ComputerName = $Computer.ComputerName.Trim()
    if ([string]::IsNullOrWhiteSpace($ComputerName)) {
        Write-Host "Computer Name is empty or null in the CSV. Skipping..."
        continue
    }

    $TotalCount++

    try {
        $ADComputer = Get-ADComputer -Identity $ComputerName -ErrorAction Stop
        $InDomainCount++
        $InDomainComputers += $ComputerName
    } catch {
        $NotInDomainCount++
    }

    $Result = [PSCustomObject]@{
        ComputerName = $ComputerName
        InDomain     = if ($InDomainComputers -contains $ComputerName) { "Yes" } else { "No" }
    }

    Write-Host "Computer Name: $($Result.ComputerName), In AD: $($Result.InDomain)"
}

Write-Host "Processing complete."
Write-Host "Total Computers Checked: $TotalCount"
Write-Host "Computers in AD: $InDomainCount"
Write-Host "Computers not in AD: $NotInDomainCount"

Write-Host "Computers in the Domain:"
$InDomainComputers | ForEach-Object { Write-Host $_ }

Write-Host "Press any key to exit..."
[System.Console]::ReadKey() > $null
