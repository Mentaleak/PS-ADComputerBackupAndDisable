# Load the System.Windows.Forms assembly to use Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Set the number of days since last logon to filter computers
$daysSinceLogon = 365
# Calculate the date $daysSinceLogon days ago
$howOLD = (get-date).AddDays(-1 * ($daysSinceLogon))
# Set the distinguished name of the OU for disabled computers
$DisabledComputersOU = 'OU=Disabled Computers,DC=Domain,DC=COM'

# Generate a timestamp for the file and OU name
$dt = "$(get-date -Format "yyyyMMdd-hhmmss")"

# Create a SaveFileDialog to choose where to save the output file
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.InitialDirectory = "C:\"
$saveFileDialog.Filter = "clixml (*.clixml)|*.clixml|All Files (*.*)|*.*"
$saveFileDialog.FileName = "DisabledComputerAccounts_$($dt).clixml"

# Display progress of the script
Write-Progress -Activity "Processing" -Status "0% Complete:" -PercentComplete 0 -CurrentOperation "Getting Computers that Have not Logged in in days"

# Get Windows computers with specific properties from Active Directory
$winComputers = Get-ADComputer -Filter {OperatingSystem -like "*Windows*"} -properties whenCreated,whenChanged,PasswordLastSet,modifyTimeStamp,samaccountname,ms-Mcs-AdmPwdExpirationTime,OperatingSystem,IPV4Address, lastlogondate, enabled, ms-Mcs-AdmPwd, distinguishedname | select samaccountname,
    @{Name="whenCreated";Expression={$_.whenCreated.ToString("yyyy-MM-dd")}},
    @{Name="whenChanged";Expression={$_.whenChanged.ToString("yyyy-MM-dd")}},
    @{Name="PasswordLastSet";Expression={$_.PasswordLastSet.ToString("yyyy-MM-dd")}},
    @{Name="modifyTimeStamp";Expression={$_.modifyTimeStamp.ToString("yyyy-MM-dd")}},
    @{Name="LAPSExpirationTime";Expression={
        $expirationTimeTicks = $_."ms-Mcs-AdmPwdExpirationTime"
        $epochStart = New-Object DateTime 1601, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
        $dateTime = $epochStart.AddTicks($expirationTimeTicks)
        if($dateTime.Year -gt 1900){
            $dateTime.ToString("yyyy-MM-dd")
        } else {
            ""
        }
    }},
    OperatingSystem,
    IPV4Address, lastlogondate, enabled, ms-Mcs-AdmPwd, distinguishedname

# Filter out computers already in the Disabled Computers OU
$winComputers = $winComputers | where {$_.DistinguishedName -notlike "*,OU=Disabled Computers,*"}

# Update progress status
Write-Progress -Activity "Processing" -Status "10% Complete:" -PercentComplete 10 -CurrentOperation "Choosing Machines to Disable"

# Filter computers that have not changed, had their password set, or LAPS expired in $howOLD days
$winComputersOLD = $winComputers | Where-Object {
    (Get-Date $_.whenChanged) -lt $howOLD -and
    (Get-Date $_.PasswordLastSet) -lt $howOLD -and
    (
        -not $_.LAPSExpirationTime -or
        (Get-Date $_.LAPSExpirationTime) -lt $howOLD
    )
}

# Initialize an array to store computers to be Backed Up
$BackedUp = @()
# Display a grid view for user to choose machines to backup data
$BackedUp += $winComputersOLD | Out-GridView -OutputMode Multiple -Title "Choose Machines to backup data"

if($BackedUp.count -gt 0){
    # Update progress status
    Write-Progress -Activity "Processing" -Status "$i% Complete:" -PercentComplete 20 -CurrentOperation "Getting BitLocker Data"
    
    # Get BitLocker recovery information from Active Directory
    $BitlockerData = Get-ADObject -Filter 'objectClass -eq "msFVE-RecoveryInformation"' -Properties *

    $count = 0
    # Add BitLocker data to each selected computer
    foreach($TermComp in $BackedUp){
        $TermComp | Add-Member -MemberType NoteProperty -Name "BitlockerData" -Value ($BitlockerData | Where-Object { $_.'DistinguishedName' -match $TermComp.DistinguishedName }) -Force
        $Count++
        # Update progress status
        Write-Progress -Activity "Processing" -Status "$i% Complete:" -PercentComplete $(($count/$BackedUp.count)*100) -CurrentOperation "Merging Bitlocker Data"
    }

    # Show save file dialog
    $saveFileDialog.ShowDialog()
    # Get the selected file name from the save file dialog
    $BackupfileName = $saveFileDialog.FileName
    
    # Export selected computers with BitLocker and Laps data to a clixml file
    $BackedUp | select Samaccountname, distinguishedname, ms-Mcs-AdmPwd, BitlockerData | Export-Clixml -Path $BackupfileName
    
    $BackupImport = @()
    # Import the clixml file and display a grid view for user to choose machines to disable
    $BackupImport += import-Clixml -Path $BackupfileName | Out-GridView -OutputMode Multiple -Title "Choose Machines to Disable"
    
    if($BackupImport.count -gt 0){
        # Create a new Organizational Unit (OU) with the current date as the name
        New-ADOrganizationalUnit -Name $dt -Path $DisabledComputersOU
        $OutOU = "OU=$($dt),$($DisabledComputersOU)"
        
        # Move the selected computers to the new OU and disable their accounts
        foreach($computer in $BackupImport){
            Disable-ADAccount -Identity $computer.DistinguishedName
            Move-ADObject -Identity $computer.DistinguishedName -TargetPath $OutOU
        }
    }
}
