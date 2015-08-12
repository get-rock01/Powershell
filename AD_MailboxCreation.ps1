# Static Entries
$ou = "OU=Users,OU=Accounts,DC=hq,DC=crabel,DC=com"
$intdomain = "hq.crabel.com"

# Data Entry section
$fname = Read-Host "First Name"
$lname = Read-Host "Last Name"
$username = Read-Host "User Account Name"
$password = Read-Host "Password"
$desc = Read-Host "Description"
$office = Read-Host "Office - MKE or LA"
$title = Read-Host "Job Title"
$mbdb = Read-Host "Which mailbox database? CH1-MB01 or LA3-MB01"
#$group = Read-Host "Name of AD group you want the user to be added to"


# Create user and enable mailbox
write-host -foregroundcolor cyan "Creating and enabling $fname $lname in AD and Exchange 2013"

pause

$pwd = convertto-securestring $password -asplaintext -force
$name = $fname + " " + $lname
$upn = $username + "@" + $intdomain
$alias = $username
$sam = $username
New-Mailbox -name $name -userprincipalname $upn -Alias $alias -DisplayName $name -OrganizationalUnit $ou -SamAccountName $sam -FirstName $fname -LastName $lname -Password $pwd –Database $mbdb


# Pause 10 for AD changes
write-host -foregroundcolor cyan "Pausing 45 Seconds for AD Changes"

pause

Start-Sleep -s 45

write-host -foregroundcolor cyan "Confirming AD account creation"

pause

$user = $username
get-aduser $user


#Disabling ActiveSync for new user
write-host -ForegroundColor Cyan "Disabling ActiveSync for new user in accordance to company policy"
Set-CASMailbox -Identity $user -ActiveSyncEnabled $false
write-host -ForegroundColor Cyan "Verifying ActiveSync is disabled for new user..."
Get-CASMailbox -Identity $user | ft Name,ActiveSyncEnabled -AutoSize

pause

#Enable Auditing on mailbox
write-host -foregroundcolor cyan "Checking mailbox auditing is enabled..."
$Audit = get-mailbox -Identity $user | select AuditEnabled

IF ($Audit -match 'True') {
    write-host -ForegroundColor cyan "Mailbox auditing is enabled on $user"
    }
ELSE {
    write-host -ForegroundColor Red "Mailbox auditing not set on $user"
    write-host -ForegroundColor Cyan "Enabling auditing now..."
    set-mailbox -Identity $user -AuditEnabled $true
	write-host -ForegroundColor Cyan "Re-checking mailbox auditing on $user"
    get-mailbox -Identity $user | ft Name,AuditEnabled -AutoSize
}

pause

# Create home directory for user
write-host -ForegroundColor Cyan "Creating home directory for $fname $lname"

pause

$teststring = "\\hq\shares\homedir\$user"
new-item -name $user -ItemType Directory -path \\hq\shares\HomeDir | out-null


# Confirm creation of user's home directory
write-host -ForegroundColor Cyan "Confirm creation of home directory"

pause

If (Test-Path $teststring) {
write-host -ForegroundColor Cyan "Home directory created!"
} 
Else {
write-host -ForegroundColor Red "Home directory does not exist"
}

pause

# Setting permissions for user's new home directory
write-host -ForegroundColor Cyan "Setting permissions on $fname $lname's home directory"

pause

$ACL = Get-Acl "\\hq\shares\homedir\$user"
$ACL.SetAccessRuleProtection($false, $false)
$rights = "hq.crabel.com\$user","FullControl", "ContainerInherit, ObjectInherit", "None", "Allow"
$accessrule =  New-Object System.Security.AccessControl.FileSystemAccessRule $rights

$acl.SetAccessRule($accessrule)
$ACL | Set-Acl


# Adding user H drive, description, job title and office location to user properties
write-host -foregroundcolor cyan "Setting additional user account parameters"

pause

get-aduser $username | %  {set-aduser $_ -HomeDrive "H:" -HomeDirectory ('\\hq\shares\homedir\' + $_.SamAccountName)}
set-aduser $username -title $title -Description $desc -Office $office


#"Account and mailbox creation for " + $fname + " " + $lname + " is complete!"
write-host -foregroundcolor cyan "Account and mailbox creation for $fname $lname is complete!"