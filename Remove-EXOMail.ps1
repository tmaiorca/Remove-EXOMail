<#............................................................................................................................................................................................ 
Purpose: Searches for email based on KQL query syntax and prompts to proceed with purge. 
Developed By: Maiorca, Troy 
Last Updated: 4/1/21 
............................................................................................................................................................................................#> 

#Module prerequisite check & download using PSGallery 
$modules = "ExchangeOnlineManagement" 

Write-Host 'Checking for prerequisite module...' -ForegroundColor Yellow 
$modules | ForEach-Object { 
if (Get-Module -ListAvailable -Name $_) { 
    Write-Host "$_ - installed" -ForegroundColor Green  
} 

else { 
    Set-PSRepository PSGallery -InstallationPolicy Trusted 
    $PSGalleryCheck = Find-Module $_ 
    Write-Host `n"$_ - not installed" -ForegroundColor Red         
    Write-Host "Downloading $_ from PSGallery..." -ForegroundColor Yellow -NoNewline 
       
    $PSGalleryCheck | Install-Module 
        if (Get-Module -ListAvailable $_) { 
            Write-Host `n"$_ module installed successfully!" -ForegroundColor Green 
            Get-Module -ListAvailable $_ | Select-Object Name, Version, ModuleType, Path 
            Set-PSRepository PSGallery -InstallationPolicy Untrusted 
        } 
        else { 
            Write-Host `n"$_ module installation failed. Please install module and re-run." 
            Exit 
        } 
    }            
} 


Write-Host `n`n'---Automated Email Deletion Script---' -ForegroundColor Green 
Write-Host 'Any emails purged from this script will be placed in the users recoverable items location for easy recover.' -ForegroundColor Yellow 

#Connecting to Exchange Online using modern authentication (to find list of recipients first using message trace) 
Write-Host `n`n'Connecting to Exchange Online to perform a message trace in order to receive recipient list...' -ForegroundColor Yellow 
Connect-ExchangeOnline 

#Defining search parameters for new eDiscovery / Compliance Search | For more information on Keywork Query Language Syntax (KQL): https://docs.microsoft.com/en-us/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference 
Write-Host 'Proceeding to gather Message Trace & Compliance Search criteria...' -ForegroundColor Yellow 
#Prompt: FROM 
$from = Read-Host "SENDER" 
#Prompt: SUBJECT 
$subject = Read-Host "SUBJECT (optional)" 
#Prompt: DATE 
Write-Host 'DATE:' 
Write-Host "1. Today's date (starting at 12:00AM)" -Foregroundcolor Yellow 
Write-Host "2. Manually specify the date (must be within 10 days of today's date)" -Foregroundcolor Yellow 
Write-Host "3. Last 7 days" -Foregroundcolor Yellow 
$choice1 = Read-Host "Please select one of the options above. The search will be done starting at 12:00AM on the respective day" 
    #Error checking - forcing the user to select 1, 2, 3 
    while ("1","2","3" -notcontains $choice1) { 
        $choice1 = Read-Host "Incorrect selection! Please select one of the options above" 
    } 
    #Option 1 - Today's date 
    if ($choice1 -eq '1') { 
        $startDate = (Get-Date -Hour 0 -Minute 0 -Second 0).ToString("MM/dd/yyyy") 
        $endDate = Get-Date 
    } 
 
    #Option 2- Manually specifying date 
    elseif ($choice1 -eq '2') { 
        Write-Host `n"Specify the date in either of the following formats. The search will be done from the date entered at 12:00AM: yyyy-MM-dd or MM/dd/yyyy" -ForegroundColor Yellow 
        while($choice1 -eq '2'){ 
            $startDate = Read-Host 
            Try{ 
                $date = [datetime]$startDate 
                break 
            } 
            Catch{ 
                Write-Host 'Not a valid date format! Please re-enter using the following format: yyyy-MM-dd' -ForegroundColor Red 
            } 
        } 
        $endDate = Get-Date 
        $endDate = [datetime]$endDate 
    } 

    #Option 3 - Last 7 days 
    if ($choice1 -eq '3') { 
        $startDate = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(-7).ToString("MM/dd/yyyy") 
        $endDate = Get-Date 
    } 

#Running Message Trace based on selected option 
Write-Host `n"Performing message trace based on the following values:" -ForegroundColor Yellow 
Write-Host "SENDER: $from" 
Write-Host "SUBJECT: $subject" 
Write-Host "DATE: $startDate" 
$msgTrace = Get-MessageTrace -SenderAddress $from -StartDate $startDate -EndDate $endDate | Where {$_.Subject -like "*$subject*"} 
$recipients = ($msgTrace.RecipientAddress) | Sort -Unique 
    
    #Storing only RVCC mailboxes & sorting uniquely to remove any duplicates based on alias/primary SMTP Address 
    $recipients = ForEach ($mailbox in $recipients) { 
    try { 
        (Get-Mailbox $mailbox -ErrorAction Stop).PrimarySmtpAddress 
        } 
    catch { 
        } 
    } 
    $recipients = $recipients | Sort -Unique 
     
    #Displaying count of recipients that received matching email content 
    $count = $recipients.count 
    Write-Host `n"Total Recipients: "$count -ForegroundColor Yellow 
        #If no recipients, exit script (won't pass variable correctly into S&C Compliance Search otherwise) 
        if ($count -eq 0) { 
            Write-Host "No recipients were found based on the entered values. Exiting script!!" -ForegroundColor Red 
            Disconnect-ExchangeOnline -Confirm:$false 
            Exit 
        } 
 
#Disconnecting from Exchange Online 
Disconnect-ExchangeOnline -Confirm:$false | Out-Null 


#Connecting to Security & Compliance Center using modern authentication (searching & purging emails found from recipients found in message trace) 
Write-Host `n`n'Connecting to Security & Compliance Center to purge matched emails based on recipients found from the Message Trace...' -ForegroundColor Yellow 
Connect-IPPSSession 

#Format KQL query with selected date option 
Write-Host "You entered the following search values from the Message Trace:" -ForegroundColor Yellow 
Write-Host "SENDER: $from",`n"SUBJECT: $subject",`n"DATE: $startdate" 
Write-Host "Press 'a' to abort, or any other key to continue with Compliance Search:" -ForegroundColor Yellow -NoNewline 
    $response = Read-Host  
    $aborted = $response -eq "a" 
    #Exit script if user entered 'a' 
    if ($aborted -eq "a") { 
        Write-Host "Aborting script & disconnecting from Exchange Online!!!" -ForegroundColor Red 
        Disconnect-ExchangeOnline -Confirm:$false 
        Exit 
    } 

    #Format KQL based on Sender, Subject, and Date values 
    else { 
        $ComplianceSearchName = "Spam_Deletion" 
        New-ComplianceSearch -Name $ComplianceSearchName -ExchangeLocation $recipients -ContentMatchQuery " 
        (From:$from) AND 
        (Subject:$subject) AND 
        (Sent>=$startDate) 
        " 
    } 

 
#Start Compliance Search using formatted KQL Query & perform a status check loop until completed 
Write-Host `n'Starting Compliance Search' -NoNewline -ForegroundColor Yellow 
Start-ComplianceSearch -Identity "$ComplianceSearchName" 
While ((Get-ComplianceSearch "$ComplianceSearchName" | Select-Object -expand Status) -ne "Completed") { 
    Start-Sleep 5; write-host "." -NoNewline 
    } 


$contentCheck = "C:\temp\Spam_Message_Details_$(Get-Date -Format yyyy-MM-dd_HH-mm).csv" 
Write-Host 'Compliance Search completed!' -ForegroundColor Green 
 
#Export Compliance Search to a local file for content validation before purge action 
Get-ComplianceSearch -Identity "$ComplianceSearchName" -Resultsize Unlimited | FL > "$contentCheck" 
Write-Host `n"Exported content to the following location: $contentCheck`nPlease validate the content before proceeding with content removal." -ForegroundColor Yellow 
Start-Process $contentCheck 

 
#Final prompt before purging content returned from Compliance Search results 
Write-Host `n'Would you like to proceed with a soft purge of all results included in the exported Complinace Search?' -ForegroundColor Yellow 
Write-Host "Press 'a' to abort, or any other key to continue soft purge:" -ForegroundColor Yellow -NoNewline 
    $response = Read-Host 
    $aborted = $response -eq "a" 
    #Prompt to quit if inputted data is incorrect 
    if ($aborted -eq "a") { 
        Write-Host `n"Aborting script & disconnecting from Exchange Online!! The most recent compliance search will be removed!" -ForegroundColor Red 
        Remove-ComplianceSearch $ComplianceSearch -Confirm:$false 
        Disconnect-ExchangeOnline -Confirm:$false 
        Exit 
    } 

    #Proceed with compliance search soft purge 
    else { 
        Write-Host `n"Proceeding with soft deletion of matched content" -NoNewline -ForegroundColor Yellow 
        New-ComplianceSearchAction -SearchName "$ComplianceSearchName" -Purge -PurgeType SoftDelete -Confirm:$false | Out-Null 
    } 

#Setting Compliance Search action variable for check and continue  
Start-sleep 5 #seconds 
$ComplianceSearchAction = $ComplianceSearchName + "_Purge" 
While ((Get-ComplianceSearchAction $ComplianceSearchAction | Select-Object -expand Status) -ne "Completed") { 
    Start-Sleep 5; write-host "." -NoNewline 
    } 

#Remvoing Compliance Search & Compliance Search Action 
Write-Host 'Soft delete of matched content completed!' -ForegroundColor Green 
Remove-ComplianceSearchaction -identity $ComplianceSearchAction -Confirm:$false 
Remove-ComplianceSearch -identity $ComplianceSearchName -Confirm:$false 

#Disconnecting from Security & Compliance Center 
Write-Host `n"" 
Disconnect-ExchangeOnline -Confirm:$false 
 
#Prompt to exit script 
Write-Host `n"Disconnected from Exchange Online & removed the Compliance Search. Press any key to exit." -ForegroundColor Yellow -NoNewline 
$response = Read-Host 
if ($response){ 
    Exit 
    }
