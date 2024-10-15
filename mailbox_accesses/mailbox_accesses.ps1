# Set the path for the export Excel file
$exportPath = "C:\Users\emerson\Downloads\projects\MailboxPermissions.xlsx"

# Install the necessary modules
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
}
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Import the modules
Import-Module ExchangeOnlineManagement
Import-Module ImportExcel

# Connect to Exchange Online
Connect-ExchangeOnline -ShowProgress $true

# Retrieve all mailboxes and permissions
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Create an array to hold the results
$mailboxPermissions = @()

foreach ($mailbox in $mailboxes) {
    try {
        # Get mailbox permissions for the specified user
        $permissions = Get-MailboxPermission -Identity $mailbox.Identity
    
        foreach ($permission in $permissions) {
            $mailboxPermissions += [PSCustomObject]@{
                Mailbox      = $mailbox.PrimarySmtpAddress
                User         = $permission.User
                AccessRights = $permission.AccessRights -join ', '
                MailboxPermission = $true
            }
        }
    }
    catch {
        # Print a message if Get-MailboxPermission fails
        Write-Host "Could not retrieve permissions for mailbox: $($mailbox.PrimarySmtpAddress). Error: $_"

        # Add a record with the MailboxPermission attribute set to false
        $mailboxPermissions += [PSCustomObject]@{
            Mailbox         = $mailbox.PrimarySmtpAddress
            User            = ""  # No user info available
            AccessRights    = ""  # No access rights available
            MailboxPermission = $false # Indicate permissions were not retrieved
        }
    }
}

# Export the results to an Excel file
$mailboxPermissions | Export-Excel -Path $exportPath -AutoSize

# Notify user of the export
Write-Host "FINISHED!!! Mailbox permissions saved to $exportPath"

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
