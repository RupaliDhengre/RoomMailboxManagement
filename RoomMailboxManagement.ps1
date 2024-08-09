# Define the menu
function Show-Menu {
    Clear-Host
    Write-Host "------------- Room Mailbox Creation -----------"
    Write-Host "1. Connect to the Exchange Online Module"
    Write-Host "2. Import the CSV file"
    Write-Host "3. Create the Room MailBoxes"
    Write-Host "4. Set the Capacity and location details"
    Write-Host "5. Validate the data"
    Write-Host "6. Exit"
}

# Function to connect to Exchange Online
function Connect-ExchangeOnline {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -UserPrincipalName user@domain.com -ShowProgress $true
}

# Function to import CSV file
function Import-CSVFile {
    Write-Host "Importing CSV file..."
    $global:roomMailboxes = Import-Csv -Path "RoomMailboxList.csv"
}

# Function to create room mailboxes
function Create-RoomMailboxes {
    Write-Host "Creating Room Mailboxes..."
    foreach ($room in $roomMailboxes) {
        New-Mailbox -Room -Name $room.DisplayName -PrimarySmtpAddress $room.Email
    }
}

# Function to set capacity and location details
function Set-CapacityAndLocation {
    Write-Host "Setting Capacity and Location details..."
    foreach ($room in $roomMailboxes) {
        Set-Mailbox -Identity $room.Email -ResourceCapacity $room.Capacity
        Set-Mailbox -Identity $room.Email -Office $room.Location
    }
}

# Function to validate the data
function Validate-Data {
    Write-Host "Validating data..."
    foreach ($room in $roomMailboxes) {
        Get-Mailbox -Identity $room.Email | Format-List DisplayName,PrimarySmtpAddress,ResourceCapacity,Office
    }
}

# Main script
do {
    Show-Menu
    $choice = Read-Host "Enter your choice"
    switch ($choice) {
        1 { Connect-ExchangeOnline }
        2 { Import-CSVFile }
        3 { Create-RoomMailboxes }
        4 { Set-CapacityAndLocation }
        5 { Validate-Data }
        6 { Write-Host "Exiting..."; break }
        default { Write-Host "Invalid choice, please try again." }
    }
    Pause
} while ($choice -ne 6)
