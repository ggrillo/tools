#Command line params
param (
    [int]$MaxDelete,
    [DateTime]$BeforeDate,
    [DateTime]$AfterDate,
    [bool]$Commit = $False,
    [bool]$Confirm = $True,
    [String]$Output = $null,
    [String]$Config = "imap-config.json"
)

function Log {
    param (
        [string]$message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    if (-not [string]::IsNullOrEmpty($Output)) {
        $message | Out-File -FilePath $Output -Append -Encoding utf8
    } else {
        Write-Output "[$timestamp] $message"
    }   
}

# returns a valid email server connected IMAP object
function InitImap {
    param(
        [string]$configFile = $null
    )

    $jsonFilePath = $configFile
    
    if (-not (Test-Path $jsonFilePath)) {
        Log "JSON file not found at path: $jsonFilePath"
        exit
    }

    $jsonContent = Get-Content -Path $jsonFilePath -Raw
    $imapConfig = $null
    try {
        $imapConfig = $jsonContent | ConvertFrom-Json
    } catch {
        Log "Failed to parse $jsonFilePath file: $($_.Exception.Message)"
        exit
    }

    # Load the Chilkat .NET assembly dll (download install from https://www.chilkatsoft.com/downloads_2_0.asp) 
    Add-Type -Path $imapConfig.ImapDll

    # Create an IMAP object
    $imap = New-Object Chilkat.Imap

    $imap.Port = $imapConfig.Port
    $imap.Ssl = $true

    # Connect to Mail Server
    if (-not $imap.Connect($imapConfig.ImapServer)) {
        Log "Failed to connect: $($imap.LastErrorText)"
        exit
    }

    # Authenticate with mail server, the password is an app password generated from the user account security settings
    if (-not $imap.Login($imapConfig.UserEmail, $imapConfig.AppPassword)) {
        Log "Login failed: $($imap.LastErrorText)"
        exit
    }

    # Select the Inbox folder
    if (-not $imap.SelectMailbox($imapConfig.Mailbox)) {
        Log "Failed to select mailbox: $($imap.LastErrorText)"
        exit
    }    

    return $imap
}

function FindEmails {
    param (
        [Chilkat.Imap]$imap,
        [DateTime]$startDate,
        [DateTime]$beforeDate
    )
    $start = $startDate.ToString('dd-MMM-yyyy')
    $end = $beforeDate.ToString('dd-MMM-yyyy')

    $searchCriteria = "SENTSINCE $start SENTBEFORE $end"
    $messageSet = $imap.Search($searchCriteria, $true)

    if ($imap.LastMethodSuccess -eq $false) {
        $($imap.LastErrorText)
        exit
    }

    $bundle = $imap.FetchHeaders($messageSet)
    if ($imap.LastMethodSuccess -eq $false) {
        $($imap.LastErrorText)
        exit
    }

    if($bundle.MessageCount -eq 0) {
        Log "No emails found between $($startDate.ToString('MM/dd/yyyy')) and $($beforeDate.ToString('MM/dd/yyyy'))"
        exit
    }

    if($bundle.MessageCount -ge 1000 -and $bundle.MessageCount -lt $MaxDelete) {
        if($Confirm) {
            $confirmMax = Read-Host "At least $($bundle.MessageCount) emails found between $($startDate.ToString('MM/dd/yyyy')) and $($beforeDate.ToString('MM/dd/yyyy')) which is < MaxDelete param [$MaxDelete], Continue [Y/N]?"
            if ($confirmMax -ne 'Y') {
                Log "Delete emails cancelled."
                exit
            }
        }   
    }

    return $bundle
}

# Imap searches return a max of 1000 emails, this function resets the imap object and searches again
function ResetImapEmails {
    param (
        [Chilkat.Imap]$imap,
        [DateTime]$AfterDate,
        [DateTime]$BeforeDate 
    )

    $imap.Disconnect()
    Start-Sleep -Seconds 10
    $imap = InitImap      
    $emails = FindEmails -imap $imap -startDate $AfterDate -beforeDate $BeforeDate    
    return $imap, $emails
}

# Validate Command Line Parameters
if ($BeforeDate -eq $null) {
    Write-Output "BeforeData parameter is required (ie -BeforeDate MM/dd/yyyy)"
    exit
}

if ($BeforeDate -ge (Get-Date)) {
    Write-Output "The date must be before today."
    exit
}

$BeforeDate = $BeforeDate.Date.AddDays(1).AddSeconds(-1)

if($AfterDate -eq $null) {
    $AfterDate = $BeforeDate.AddYears(-1)  # Default to 1 year 
}

if($MaxDelete -eq 0) {
    $MaxDelete = 1000  # this is the max that can be deleted at a time
}

$afterDateStr = $($AfterDate.ToString("MM/dd/yyyy")) 
$beforeDateStr = $($BeforeDate.ToString("MM/dd/yyyy")) 

# command line param should be True: when running the script manually, False: when run from a scheduled task/external script
if ($Confirm) {
    $confirmation = Read-Host "Delete $MaxDelete emails between $afterDateStr and $beforeDateStr, Continue (Y/N)?"
    if ($confirmation -ne 'Y') {
        Log "Delete emails cancelled."
        exit
    }
}
Write-Output "Searching for emails between $afterDateStr and $beforeDateStr to delete..."

# Begin Main script execution 
Log "Connecting to the IMAP mail server using Config $Config"
$imap = InitImap $Config

Log "Finding all emails between $afterDateStr and $beforeDateStr"
$emails = FindEmails -imap $imap -startDate $AfterDate -beforeDate $BeforeDate
Log "Found $($emails.MessageCount) emails between $afterDateStr and $beforeDateStr"

$emailsDeleted = 0
$errors = 0
$emailsRead = 0
$restartTries = 0
$maxImapSearchResults = 1000
$i = 0
$iteration = 1
Log "Running iteration $iteration"

while ($i++ -le $emails.MessageCount) {
    try {
        $email = $emails.GetEmail($i)        
        $emailsRead++
        $receivedDate = [DateTime]::Parse($email.EmailDate)
        
        Log "DELETING EMAIL SENDER: $($email.FromAddress), SUBJECT: $($email.Subject), RECEIVED: $($receivedDate.ToString("MM/dd/yyyy"))"
        
        # Mark the email for deletion
        [void]$imap.SetMailFlag($email, "Deleted", 1)

        $emailsDeleted++   
        if ($emailsDeleted -ge $MaxDelete) {
            Log "Reached max emails to delete ($MaxDelete)..."
            break
        }  
    } catch {
        $errors++
        Log "An error occurred: $($_.Exception.Message)"
        Log "Failed to fetch or delete email, reason: $($imap.LastErrorText)"

        # Check if the number of errors exceeds the threshold (5) and restart if necessary
        if($errors -gt 5) {
            if($restartTries -lt 3) {
                $restartTries++            
                Log "Due to errors, attempting to reconnect to the imap mail server and restart..."  
                $imap, $emails = ResetImapEmails -imap $imap -AfterDate $AfterDate -BeforeDate $BeforeDate              
                $i = 0
            } else {  
                Log "Too many errors, max restart attempts reached, exiting..."
                exit  
            }
        }
    } finally {
        if ($i -eq $emails.MessageCount -and $mails.MessageCount -eq $maxImapSearchResults) {
            Log "Max imap search emails reached [$maxImapSearchResults], resetting imap mail server connection and continuing to delete requested emails ..." 
            Log "Running iteration $iteration++" 
            $imap, $emails = ResetImapEmails -imap $imap -AfterDate $AfterDate -BeforeDate $BeforeDate              
            $i = 0
        }
    }
}

# Permanently remove the emails marked for deletion
if ($Commit) {
    Log "Permanently DELETING emails from the server..."
    if (-not $imap.Expunge()) {
        Log "Failed to permanently delete emails, reason: $($imap.LastErrorText)"
        exit
    }
    Log "Summary Emails Read: $emailsRead, Deleted: $emailsDeleted, Errors: $errors"    
} else {
    Log "-Commit parameter is false, emails marked for Deletion: $($emailsRead - $errors)"
}

# Disconnect from the IMAP server
[void]$imap.Disconnect()
