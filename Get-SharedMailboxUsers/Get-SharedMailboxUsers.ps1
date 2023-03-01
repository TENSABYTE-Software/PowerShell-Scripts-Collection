[CmdletBinding()]
param(
    [Parameter(Mandatory=$false,Position=0)]
    [string]$User = "*",
    [Parameter(Mandatory=$false)]
    [switch]$ExportCSV
)

# Check if ExchangeOnlineManagement is installed on the PC
$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement
if (!$module) {
    Write-Warning "The ExchangeOnlineManagement module is not installed on this PC. `
    Attempting to install...."
    try {
        Install-Module -Name ExchangeOnlineManagement -Force
    }
    catch {
        Write-Error "ExchangeOnlineManagement Not Installed, Install Attempt Failed. ` 
            Please make sure you are running PowerShell as an Administrator"
        return
    }
    
}


# Connect to Exchange Online
Connect-ExchangeOnline


# param splatting for Get-Mailbox

$mb_params = @{
    'RecipientTypeDetails' = 'SharedMailbox';
    'ResultSize' = 'Unlimited';
    'Filter' = '{RecipientTypeDetails -eq "SharedMailbox"}'
}

# Get all shared mailboxes
if ($User -eq "*") {
    $n = '_all_users'
    $sharedMailboxes = Get-Mailbox @mb_params | `
    %{Get-MailboxPermission $_.SamAccountName} | `
    Where-Object {$_.User -notlike 'NT AUTHORITY\SELF'}
}
else {
    $n = $User
    $sharedMailboxes = Get-Mailbox @mb_params | ` 
    Where-Object { (Get-MailboxPermission $_.Alias).User -like $User }
}

# Format the output
if ($ExportCSV) {
    $sharedMailboxes | Export-Csv -NoTypeInformation -Path "shared_mailboxes$($n).csv"
}
else {
    $sharedMailboxes
}