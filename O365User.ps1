<#
------------------------------------------------
Powershell Script for O365 User creation
Email user creation script for Office365
Written by Progenyofeniac 5/28/2018
7/5/18 Added functionality to allow for MFA-enabled admin accounts, now requires two logins
Loads email user info from $CSVPath, prompts for O365 credentials, creates user mailboxes.

The CSV file must be formatted as shown here, with headers:

FirstName  MidInitial    LastName  Department  EmployeeNumber  License               AddlGroup
John       A             Smith     PlantOps    09876           Kiosk                 Org-wide
Jane       Q             Jones     Admin       09123           Exchange              Admin

Notes: Middle initial is optional. 
       Department should be chosen from a current list to standardize, but the script will use whatever text is provided.
       Employee number gets written to CustomAttribute1 since there is no standard field for it in O365.
       License should be either Kiosk or Exchange (if using Outlook). The script converts these two to O365 names.
       All users get added to Campus Wide by default but if they should be added to an additional group, provide it here. Must be a valid group name.
#>

$global:NewUserInfoList = @() # For storing a table of user data as each user is created. This is set as a global variable because it is used inside the Add-EmailAccount function.

# Set the location of the folder/filename to load

$CSVFolder = [environment]::getfolderpath("mydocuments")
$CSVPath = "$CSVFolder\NewEmailUsers.csv"
#Set-Variable -Name $NewUserInfoList -Value @() -Scope Global

function Add-EmailAccount 
{

# The user creation function
# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# $EmailAddress is set in the script prior to calling this function. The script will determine if the email address already exists
#     and it will add the middle initial if needed.
# $UserData is an array built from the imported CSV file
# $License is the O365 naming convention of the license for each user. The script converts the "License" field of the CSV file to the O365 license name.

Param ([string]$EmailAddress,[array]$UserData,[string]$License)
Process 
  {
   Write-Host "Creating account" $EmailAddress
   $NewUserInfo = New-MsolUser -DisplayName ($UserData.FirstName + " " + $UserData.LastName) -FirstName $UserData.FirstName -LastName $UserData.LastName -UserPrincipalName $EmailAddress -Department $UserData.Department -LicenseAssignment $License -UsageLocation "US"
   $MailboxExist = $false
   Write-Host -NoNewline "Waiting on mailbox creation..."
   while ($MailboxExist -eq $false) 
     {
      Write-Host -NoNewline "."
      Start-Sleep 10
      $MailboxExist = [bool](get-mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue)
     }
   Write-Host "OK."
   Write-Host "Setting CustomAttribute1 (employee number) on" $EmailAddress "to" $UserData.EmployeeNumber 
   Set-Mailbox -Identity $EmailAddress -CustomAttribute1 $UserData.EmployeeNumber
   if ($UserData.AddlGroup -ne "")
     {
      Write-Host "Adding" $EmailAddress "to" $UserData.AddlGroup "group."
      Add-DistributionGroupMember -Identity $UserData.AddlGroup -Member $EmailAddress
     }
     else 
     {
      Write-Host "No additional groups specified for" $EmailAddress "."
     }
   Write-Host "Adding" $EmailAddress "to Campus Wide group."
   Add-DistributionGroupMember -Identity "Campus Wide" -Member $EmailAddress
   $global:NewUserInfoList += $NewUserInfo | Select @{Label="Name";Expression={$_.DisplayName}},@{Label="Email Address";Expression={$_.UserPrincipalName}},Password
  }

# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# End Add-EmailAccount function
}


# Check if the file exists before loading, exit if not
if (([bool](Test-Path $CSVPath) -eq $false))
  {
   Write-Host $CSVPath "cannot be found. Please verify the file and try again."
  }
 
  # Load data from $CSVPath, prompt user to verify these are correct users in case old data was left in the file.
  
else 
  {
   Write-Host "Loading new email user data from" $CSVPath
   $NewEmailUsers = Import-Csv $CSVPath
   Write-Host "The following users will be created:"
   $NewEmailUsers | select FirstName,LastName | where {$_.FirstName -ne ""}| Format-Table -AutoSize | Out-String | %{Write-Host $_}
   $VerifyUsers = Read-Host "Is this correct? [Y/N]"
   if ($VerifyUsers -eq "Y")
     {
      Write-Host "NOTE: You will be prompted for login twice. This is normal."
      Write-Host "Ensure that you are using an admin account for this process."
      
      # Connect to O365, download all existing email addresses & display names to compare with potential new addresses    
      # Added new login setup for MFA-enabled admin accounts 7/5/18 KJ (5 lines)
      # This new setup requires two logins: one for the Exo module and another for 'connect-msolservice'
      # Each one performs different functions and both are required, therefore two logins are required

      Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName|?{$_ -notmatch "_none_"}|select -First 1) 
      $EXOSession = New-ExoPSSession 
      Import-PSSession $EXOSession
      Import-Module MSOnline
      Connect-MsolService
      
      <#
      ------ OLD LOGIN SETUP, PRIOR TO ENABLING MFA -------
      Import-Module MSOnline
      $O365Cred = Get-Credential
      $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
      Import-PSSession $O365Session 
      Connect-MsolService –Credential $O365Cred >$null 2>&1
      ------ END OF OLD LOGIN SETUP -------
      #>

      Write-Host "Getting a full list of current email addresses..."
      $EmailAddresses = Get-Mailbox -ResultSize unlimited -RecipientTypeDetails UserMailbox | select PrimarySmtpAddress,DisplayName
      # Calculate available vs needed licenses for both Kiosk and Exchange Online, prompt user to purchase more if needed.
      Write-Host "Checking license availability..."
      $UnusedKioskLicenses = Get-MsolAccountSku | where{$_.AccountSkuId -eq "OrganizationDomain:EXCHANGEDESKLESS"} | %{$_.ActiveUnits - $_.ConsumedUnits}
      $UnusedExchangeLicenses = Get-MsolAccountSku | where{$_.AccountSkuId -eq "OrganizationDomain:EXCHANGESTANDARD"} | %{$_.ActiveUnits - $_.ConsumedUnits}
      $NeededKioskLicenses = ($NewEmailUsers | Where-Object {$_.License -eq "Kiosk"}).count
      $NeededExchangeLicenses = ($NewEmailUsers | Where-Object {$_.License -eq "Exchange"}).count
      if ($UnusedKioskLicenses -lt $NeededKioskLicenses)
         {
          Write-Host "You only have" $UnusedKioskLicenses "Kiosk licenses available but you need" $NeededKioskLicenses "."
          Read-Host "Please go online and purchase more, wait 2-3 minutes, then press [Enter] to continue"
          Write-Host "Continuing script..."
         }
      if ($UnusedExchangeLicenses -lt $NeededExchangeLicenses)
         {
          Write-Host "You only have" $UnusedExchangeLicenses "Exchange licenses available but you need" $NeededExchangeLicenses "."
          Read-Host "Please go online and purchase more, wait 2-3 minutes, then press [Enter] to continue"
          Write-Host "Continuing script..."
         }

    
      # Cycle through each new user from the imported CSV data where "FirstName" is not empty. 
      # This allows for cases where the exported CSV file has empty fields.    

      foreach ($NewUser in $NewEmailUsers) 
        {
         if ($NewUser.FirstName -ne "")

         # Convert the CSV file's license data to O365 license names. CSV file uses either "Kiosk" or "Exchange".

           {
            if ($NewUser.License -eq "Exchange")
              {
               $O365License = "OrganizationDomain:EXCHANGESTANDARD"
              }
            else 
              {
               $O365License = "OrganizationDomain:EXCHANGEDESKLESS"
              }

            # Create the preferred first initial, last name email address, see if it exists in the list of current users. Add account if not.

            $NewEmailAddressA = ((($NewUser.FirstName).Substring(0,1))+($NewUser.LastName)).tolower()+"@domain.org"
            $ExistingEmailAddressA = $EmailAddresses | ?{$_.primarysmtpaddress -eq $NewEmailAddressA}
            if ($ExistingEmailAddressA -eq $null)
              {
               Add-EmailAccount -EmailAddress $NewEmailAddressA -UserData $NewUser -License $O365License
              }

            # If first initial, last name already exists see if there's a middle initial in the CSV file. Add it to the email address if so.
            # Check if that creates a unique address, create it if so, exit and inform the user if that's still not a unique address.

              else 
                {
                 if ($NewUser.MidInitial -ne "")
                   {
                    $NewEmailAddressB = ((($NewUser.FirstName).Substring(0,1))+($NewUser.MidInitial)+($NewUser.LastName)).tolower()+"@domain.org"
                    $ExistingEmailAddressB = $EmailAddresses | ?{$_.primarysmtpaddress -eq $NewEmailAddressB
                   }
                    if ($ExistingEmailAddressB -eq $null)
                      {
                       Write-Host "Email address" $NewEmailAddressA "already exists for" $ExistingEmailAddressA.DisplayName "but" $NewEmailAddressB "is being created."
                       Add-EmailAccount -EmailAddress $NewEmailAddressB -UserData $NewUser -License $O365License}
                      else 
                        {
                         Write-Host "Both" $NewEmailAddressA "and" $NewEmailAddressB "already exist. Please manually create this account."
                        }
                   }
                   else 
                     {
                      Write-Host $NewEmailAddressA "already exists for" $ExistingEmailAddressA.DisplayName ". Please provide a middle initial or manually create this account."
                     }
                }
           }

        }
      Write-Host "The following email accounts were created:"
      $NewUserInfoList | Format-Table -AutoSize
      $SingleUser = ($Mailbox | Select DisplayName,@{Label="EmailAddress";Expression={$User.PrimarySmtpAddress}},@{Label="LastLogon";Expression={$LastLogon}})
     }
     else {Write-Host "Script is exiting. Check user file before continuing."}
  }
