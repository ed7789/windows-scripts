<#

contact-manager-expiration.ps1

Based on: https://social.technet.microsoft.com/Forums/windowsserver/en-US/9d080c24-b2a2-4d9b-b50b-ca7fb9d95a91/account-expiration-email-notification?forum=winserverpowershell

Expired/Expiring Accounts Notification Script
Modified 2016.02.18 - Scrubbed for Public View
Description:  Identifies and sends notifications about expired/expiring user accounts in AD
  1) Gets list of all AD users
	2) Parses list for only managers and obtains Exchange OOF status for each manager with expiring employees
	3) Emails each manager a list ($body) of expired/expiring employees
	4) Builds master list ($summary) of all managers' expired/expiring employees
	5) Appends to master list all expired/expiring employees without a manager ($bodyNM)
	6) Emails master list to helpdesk/PMO
	7) If there are no expired/expiring accounts, an email to that effect is sent instead of a summary.

Usage:
  This script can be run in a PowerShell 3.x or higher command line. It takes no parameters.

Changelog:
  - Sometime in 2017 by EdwinG@INAP 
    Extract variables from actual code. Also, checks if accounts in a specific OU don't have expiration dates or managers.

  - 2019/05/01 by EdwinG@INAP
    Cleaned up and made source code public
#>

<# 
  BEGIN CONFIGURATION
#>

# How many days to notify before expiration
$DaysBeforeNotification = 30
$NotifyExpired = $false # Add expired users to the manager notification message
$NotifyNoExpiration = $false # Bothers IT if no one is expiring
$LogFile = "$(Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)\contact-manager-expiration.log" # Log file

# Base OU to look in
$LDAPBase = "DC=example,DC=com"

# Mail Server configuration 
$MailServer = $null # Mail server address
$MailFrom = "Source <it@example.com>" # Email address messages are coming from
$MailSummaryTo = "Recipient <it@example.com>" # Email address to send summary emails to
$MailNoticeSubject = "IMPORTANT - Accounts due to expire" # Notice email subject to managers
$MailSummarySubject = "Summary - Expiring accounts" # Summary email subject to it
$MailBcc = "" # Email address to BCC
$MailNotifySuperior = $false # Notifies second-level manager as well, if set to $true. It informs 2nd level at all times if 1st level is disabled.
<#
  END CONFIGURATION

  BEGIN TEMPLATES
#>

#Cascading Sheet Style (CSS) for Email Body Format (managerNotice/itNotice$itNotice/itNoExpireNotice)

# HTML Message header
$header = @'
<!DOCTYPE html>
<html>
  <head>
  </head>
  <body>
'@

# HTML message footer
$footer = @'
</body>
</html>
'@

# Manager notification formats/content for active 'Manager/Supervisor' Email
$managerNoticeHeader = "$($header)
<h1>User accounts set to expire</h1>
<p>
  The user(s) in the following list report to you and their accounts are about to expire.<br />
  <br />
  Thank you.
</p>"

$managerNoticeFooter = $footer.Replace('#YEAR#', (Get-Date -UFormat "%Y")) # We take the the current footer and replace the #YEAR# string by the current year

# Manager 2nd level notification formats/content for disabled 'Manager/Supervisor' Email
$managerSuperiorNoticeHeader = "$($header)
<h1>User accounts set to expire</h1>
<p>
  The user(s)) in the following list used to report to one of your former managers and their accounts are about to expire.<br />
  <br />
  Thank you.
</p>"
$managerSuperiorNoticeFooter = $footer.Replace('#YEAR#', (Get-Date -UFormat "%Y"))

# IT notification formats/content for 'Summary' Email
$itNoticeHeader = "$($header)
<h1>Summary of User Accounts Expired or About to Expire</h1>
<p>
  The following list of user accounts are expired, are about to expire or are missing manager information.<br />
</p>"
$itNoticeFooter = $footer.Replace('#YEAR#', (Get-Date -UFormat "%Y"))

# IT notification formats/content for 'No Expired/Expiring Accounts' Email
$itNoExpireNoticeHeader = "$($header)
<h1>No Accounts Found Expired or About to Expire</h1>
<p>
  There are no accounts found expired or about to expire today.
</p>"

$itNoExpireNoticeFooter = $footer.Replace('#YEAR#', (Get-Date -UFormat "%Y"))


<# 
  END TEMPLATES
#>

# Start logging
Start-Transcript -Path $LogFile

$sendSumEmail = $false

#Build array ($summary) of ALL Expired/Expiring Accounts
$summary = @()

# This will quit the script if mail server is not configured
if($MailServer -eq $null -or $MailServer.Trim() -eq "" {
  Throw "The script is not configured. Please configure it"

} 
# Connect to Active Directory
Import-Module ActiveDirectory

# Build array consisting of all enabled AD managers users
# We double filter (true and false) because the -Filter attribute is manadatory for Get-ADUser but we don't care about it actually
Get-ADUser -Filter { enabled -EQ "true" -or enabled -EQ "false" } -Properties directReports,EmailAddress,enabled,Name,Manager,description,Department | ForEach-Object {
  #Each $body array will be emailed to respective manager, then appended to $summary array
  $body = @()
  if ($_.directReports) { # If there are any direct reports
    $itmanagercc = ""
    Clear-Variable -Name itmanagercc

    # $_ refers to the manager
    $managerEmailAddress = $_.EmailAddress
    if ($_.Enabled) {
      $managerName = $_.Name
    } else {
      $managerName = "$($_.Name) (Disabled)"
    }
    $managerDescription = $_.department

    $_.Manager | ForEach-Object { # Run for every superior of the supervisor
      if ($_ -ne $null) {
        $managerSuperiorDetails = Get-ADUser $_ -Properties EmailAddress,Name,enabled
        if ($managerSuperiorDetails.Enabled) {
          $managerSuperiorName = $managerSuperiorDetails.Name
        } else {
          $managerSuperiorName = "$($managerSuperiorDetails.Name) (Disabled)"
        }
        $managerSuperiorEmail = $managerSuperiorDetails.EmailAddress
      }
    }

    $_.directReports | ForEach-Object { # This every person that reports to the manager
      $userDetails = Get-ADUser $_ -Properties AccountExpirationDate,enabled,Company,Description
      if (($userDetails.AccountExpirationDate) -and ($userDetails.Enabled -eq "True")) { # Checks if the user is enabled and has an expiration date
        if ($userDetails.AccountExpirationDate -lt (Get-Date).AddDays($DaysBeforeNotification)) { #.AddDays parameter determines threshold for notification
          $sendSumEmail = $true

          # Data for the manager notification
          $props = [ordered]@{
            'Username' = $userDetails.SamAccountName
            'Full Name' = $userDetails.Name
            'Company' = $userDetails.Company
            'Account Expiration Date' = $userDetails.AccountExpirationDate.ToString('f')
          }

          # Data for summary email
          $propsNM = [ordered]@{
            'Username' = $userDetails.UserPrincipalName
            'Full Name' = $userDetails.Name
            'Company' = $userDetails.Company
            'Description' = $userDetails.Description
            'Account Expiration Date' = $userDetails.AccountExpirationDate.ToString('f')
            'Primary Manager' = $managerName
            'Secondary Manager' = $managerSuperiorName
          }


          if (($NotifyExpired -eq $false -and $userDetails.AccountExpirationDate -ge (Get-Date)) -or ($NotifyExpired)) {
            $sendMgrEmail = $true
            $body += New-Object PsObject -Property $props
          }

          $summary += New-Object PsObject -Property $propsNM
        }
      }
    }
  }

  if ($sendMgrEmail) {
    $sendMailExtra = ""

    if (!$_.Enabled -and $managerSuperiorDetails.Enabled) { # If the manager is disabled and his superior is enabled
      $managerEmailAddress = $managerSuperiorEmail
      $body = $body | Sort-Object 'Account Expiration Date' | ConvertTo-Html -PreContent $managerSuperiorNoticeHeader -PostContent $managerSuperiorNoticeFooter | Out-String
    } else {
      $body = $body | Sort-Object 'Account Expiration Date' | ConvertTo-Html -PreContent $managerNoticeHeader -PostContent $managerNoticeFooter | Out-String
    }

    if ($MailNotifySuperior -and $managerSuperiorDetails.Enabled) { # If we want to notify the superior
      $sendMailExtra = $sendMailExtra + " -Cc '$($managerSuperiorEmail)'"
    }

    if ($MailBcc -ne $null -and $MailBcc -ne "") { # If we want to BCC someone
      $sendMailExtra = $sendMailExtra + " -Bcc '$($MailBcc)'"
    }

    $sendmailcommand = "Send-MailMessage -From '" + $MailFrom + "' -To '" + $managerEmailAddress + "' " + $sendMailExtra + " -Subject '" + $MailNoticeSubject + "'" + ' -Body $body' + " -BodyAsHTML -SmtpServer '" + $MailServer + "' -Priority High "

    # If both the manager and the superior are disabled, I can't send an email to them.
    if ($_.Enabled -or $managerSuperiorDetails.Enabled) {
      Invoke-Expression $sendmailcommand

      Write-Output "Sent email to $($managerEmailAddress)"
    } else {
      Write-Output "Can't send email to $($managerName) or $($managerSuperiorName). Both are disabled."
    }
    Clear-Variable itmanagercc
  }
  $sendMgrEmail = $false
}

# Append expired/expiring accounts with no manager to $summary to be emailed to IT in the base OU
Get-ADUser -SearchBase $LDAPBase -Filter { enabled -EQ "true" } -Properties AccountExpirationDate,Manager,Company,Description | Sort-Object AccountExpirationDate | ForEach-Object {
  if (!$_.Manager) {
    if ($_.AccountExpirationDate) {  # There is an expiration date but no manager
      if ($_.AccountExpirationDate -lt (Get-Date).AddDays($DaysBeforeNotification)) {
        $sendSumEmail = $true
        $propsNM = [ordered]@{
          'Username' = $_.UserPrincipalName
          'Full Name' = $_.Name
          'Company' = $_.Company
          'Description' = $_.Description
          'Account Expiration Date' = $_.AccountExpirationDate.ToString('f')
          'Primary Manager' = "NO MANAGER ASSIGNED"
          'Secondary Manager' = "NO MANAGER ASSIGNED"
        }
        $summary += New-Object PsObject -Property $propsNM
      }
    } else {
      # No expiration date set but no manager
      $sendSumEmail = $true
      $propsNM = [ordered]@{
        'Username' = $_.UserPrincipalName
        'Full Name' = $_.Name
        'Company' = $_.Company
        'Description' = $_.Description
        'Account Expiration Date' = "NO EXPIRATION DATE"
        'Primary Manager' = "NO MANAGER ASSIGNED"
        'Secondary Manager' = "NO MANAGER ASSIGNED"
      }
      $summary += New-Object PsObject -Property $propsNM
    }
  } else { # There is a manager, but no expiration date
    $manager = Get-ADUser $_.Manager -Properties Name,Manager,enabled
    if (!$_.AccountExpirationDate) {
      if ($manager.Enabled) {
        $managerName = $manager.Name
      } else {
        $managerName = "$($manager.Name) (Disabled)"
      }

      if ($manager.Manager -ne $null) {
        $managerSuperiorDetails = Get-ADUser $manager.Manager -Properties Name,enabled

        if ($managerSuperiorDetails.Enabled) {
          $managerSuperiorName = $managerSuperiorDetails.Name
        } else {
          $managerSuperiorName = "$($managerSuperiorDetails.Name) (Disabled)"
        }
      } else {
        $managerSuperiorName = 'NO MANAGER ASSIGNED'
      }

      $sendSumEmail = $true
      $propsNM = [ordered]@{
        'Username' = $_.UserPrincipalName
        'Full Name' = $_.Name
        'Account Expiration Date' = 'NO EXPIRATION DATE'
        'Primary Manager' = $managerName
        'Secondary Manager' = $managerSuperiorName
      }
      $summary += New-Object PsObject -Property $propsNM
    }
  }
}

$sendMailExtra = ""

if ($MailBcc -ne $null -and $MailBcc -ne "") {
  $sendMailExtra = $sendMailExtra + " -Bcc '$($MailBcc)'"
}

if ($sendSumEmail) {
  Write-Output "Sending summary email to IT"
  $summary = $summary | Sort-Object 'Account Expiration Date' | ConvertTo-Html -PreContent $itNoticeHeader -PostContent $itNoticeFooter | Out-String
  $sendmailcommand = "Send-MailMessage -From '" + $MailFrom + "' -To '" + $MailSummaryTo + "' " + $sendMailExtra + " -Subject '" + $MailSummarySubject + "'" + ' -Body $summary' + " -BodyAsHTML -SmtpServer '" + $MailServer + "'"
  Invoke-Expression $sendmailcommand
}
else {
  Write-Output "No one expires today."
  $noexpired = ConvertTo-Html -PreContent $itNoExpireNoticeHeader -PostContent $itNoExpireNoticeFooter | Out-String

  if ($NotifyNoExpiration) {
    $sendmailcommand = "Send-MailMessage -From '" + $MailFrom + "' -To '" + $MailSummaryTo + "' " + $sendMailExtra + " -Subject '" + $MailSummarySubject + "'" + ' -Body $summary' + " -BodyAsHTML -SmtpServer '" + $MailServer + "'"
    Invoke-Expression $sendmailcommand
  }
}

# Stop logging
Stop-Transcript