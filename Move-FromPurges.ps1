﻿# Move-FromPurges.ps1
# Last Edited: 9/13/2017
# By Brad
#
# DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# The Author FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# The Author OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF The Author HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

param(
	[Parameter(Position=0,Mandatory=$True,HelpMessage="Specifies the mailbox(es) to be accessed from an array (use comma seperate)")]
	[ValidateNotNullOrEmpty()]
    $Mailboxes,
	[Parameter(Position=1,Mandatory=$false,HelpMessage="Specifies the subfolder under Inbox that you want the messages moved to")]
	[ValidateNotNullOrEmpty()]
    [string]$subfolder,
	[Parameter(Position=2,Mandatory=$True,HelpMessage="Specify the UPN of the account that has impersonation rights to the mailboxes you've specified")]
	[ValidateNotNullOrEmpty()]
    [string]$AccountWithImpersonationRights, 
    [Parameter(Position=3,Mandatory=$True,HelpMessage="Specify Start Date Format YYYY-MM-DD")]
    [ValidateNotNullOrEmpty()]
    [string]$startdate,
    [Parameter(Position=4,Mandatory=$True,HelpMessage="Specify End Date YYYY-MM-DD")]
    [ValidateNotNullOrEmpty()]  
    [string]$enddate,
    [Parameter(Position=5,Mandatory=$false,HelpMessage="Use -Whatif to show what the script would do")]
    [ValidateNotNullOrEmpty()]
    [switch]$whatif,
    [Parameter(Position=5,Mandatory=$false,HelpMessage="Default PageLimit 1000 (this is the maximum), specify integer value lower than maximum")]
    [ValidateNotNullOrEmpty()]
    [int]$pagelimit,
    [Parameter(Position=6,Mandatory=$false,HelpMessage="This will recover tasks to the tasks folder, it is a switch parameter")]
    [ValidateNotNullOrEmpty()]
    [switch]$recoverTasks,
    [Parameter(Position=7,Mandatory=$false,HelpMessage="This will recover Notes to the Notes folder, it is a switch parameter")]
    [ValidateNotNullOrEmpty()]
    [switch]$recoverNotes,
    [Parameter(Position=8,Mandatory=$false,HelpMessage="This will recover Calendar items to the Calendar folder, it is a switch parameter")]
    [ValidateNotNullOrEmpty()]
    [switch]$recoverCalendar,
    [Parameter(Position=8,Mandatory=$false,HelpMessage="This will recover all items, it is a switch parameter")]
    [ValidateNotNullOrEmpty()]
    [switch]$recoverAll,
    [string]$logpath,
    [int]$threadlimit      
    )
if(!$logpath){
    $rootpath = "$($PWD.Path)\"
    $logpath = $rootpath + "Purge-Recovery-Master.txt"
}
else{
    $rootpath = $rootpath.TrimEnd("\") + "\"
}
[array]$Mailboxes = $Mailboxes
#check for base components:
$dllexists = Get-ChildItem "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" -ErrorAction SilentlyContinue
if(!$dllexists){Write-Host "No Microsoft.Exchange.WebServices.dll found in C:\Program Files\Microsoft\Exchange\Web Services\2.2\ ; Cannot Continue, Please Install the EWS Manage API 2.2"; exit;}
$poshcmdexist = Get-Command -Module PoshRSJob -ErrorAction SilentlyContinue
if(!$poshcmdexist){Write-Host "Missing PoshRsJob Module, this module is necessary for multithreading; You can install with Install-Module PoshRSJob from an Administrative Powershell window (PowershellGet required for Powervshell v3 and lower)"; exit;}
Import-Module PoshRsJob
#Credential Management
Get-RSJob -State Completed | Remove-RSJob
$psCred = Get-Credential -UserName $AccountWithImpersonationRights -Message "Enter Impersonation Account Credentials"
Function Write-LogEntry {

    param(
        [string] $LogName ,
        [string] $LogEntryText,
        [string] $ForegroundColor
    )
    if ($LogName -NotLike $Null) {
        # log the date and time in the text file along with the data passed
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : $LogEntryText" | Out-File -FilePath $LogName -append;
        if ($ForeGroundColor -NotLike $null) {
            # for testing i pass the ForegroundColor parameter to act as a switch to also write to the shell console
            write-host $LogEntryText -ForegroundColor $ForeGroundColor
        }
    }
}
if($Mailboxes[0] -like "$*"){
    [string]$usevar = ($Mailboxes[0]).Trim("$")
    Write-LogEntry -LogName $logpath -LogEntryText "Trying to convert bad input: $usevar"
    [array]$Mailboxes = @(Get-Variable -Name $usevar -ValueOnly)
}
Write-LogEntry -LogName $logpath -LogEntryText "Users Being Migrated:"
Write-LogEntry -LogName $logpath -LogEntryText "$($mailboxes)"

$scriptblock = {

    $params = $Using:paramblock
    $creds = $Using:psCred
    #[Parameter(Position=0)]
    [string]$MailboxToImpersonate = $params.MailboxToImpersonate
    #[Parameter(Position=1)]
    [string]$subfolder = $params.subfolder
    #[Parameter(Position=2)]
    [string]$AccountWithImpersonationRights = $params.AccountWithImpersonationRights
    #[Parameter(Position=3)]
    [datetime]$startdate = $params.startdate
    #[Parameter(Position=4)]
    [datetime]$enddate = $params.enddate
   # [Parameter(Position=5)]
    [switch]$whatif = $params.whatif
    #[Parameter(Position=6)]
    [int]$pagelimit = $params.pagelimit
    #[Parameter(Position=7)]
    [string]$logpath = $params.logpath
    #[Parameter(Position=8)]
    #[System.Management.Automation.PSCredential]$psCred   

    $Error.Clear()
    $logpath = $logpath + "Purge-Recovery-$MailboxToImpersonate.txt"
    Function Write-LogEntry {
        param(
            [string] $LogName ,
            [string] $LogEntryText,
            [string] $ForegroundColor
        )
        if ($LogName -NotLike $Null) {
            # log the date and time in the text file along with the data passed
            "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : $LogEntryText" | Out-File -FilePath $LogName -append;
            if ($ForeGroundColor -NotLike $null) {
                # for testing i pass the ForegroundColor parameter to act as a switch to also write to the shell console
                write-host $LogEntryText -ForegroundColor $ForeGroundColor
            }
        }
    }
    Write-LogEntry -LogName $logpath -LogEntryText "Starting Recover for: $mailboxToImpersonate on $(Get-Date)"
    [hashtable]$Global:recoverOptions = @{}
    if(!($params.recoverTasks) -and !($params.recoverCalendar) -and !($params.recoverNotes) -and !($params.recoverAll)){
        Write-LogEntry -logName $logpath -LogEntryText "Detected no Recovery Options, setting RecoverEmail to True"
        $Global:recoverOptions.RecoverEmail = "Yes"
        Write-LogEntry -LogName $logpath -LogEntryText "$($Global:recoverOptions)"
    }
    elseif($params.recoverAll -eq "Yes"){
        Write-LogEntry -LogName $logpath -LogEntryText "Detected RecoverAll, setting all options to True"        
        $Global:recoverOptions.recoverTasks = "Yes"
        $Global:recoverOptions.recoverNotes = "Yes"
        $Global:recoverOptions.recoverCalendar = "Yes"
        $Global:recoverOptions.RecoverEmail = "Yes"
        Write-LogEntry -LogName $logpath -LogEntryText "$($Global:recoverOptions)"        
    }
    else{
        Write-LogEntry -LogName $logpath -LogEntryText "Setting options to individual preferences"
        if($params.RecoverEmail){$Global:recoverOptions.RecoverEmail = $params.RecoverEmail}
        if($params.recoverCalendar){$Global:recoverOptions.recoverCalendar = $params.recoverCalendar}
        if($params.recoverNotes){$Global:recoverOptions.recoverNotes = $params.recoverNotes}
        if($params.recoverTasks){$Global:recoverOptions.recoverTasks = $params.recoverTasks}
        Write-LogEntry -LogName $logpath -LogEntryText "$($Global:recoverOptions)"        
    }

    if(!$creds){throw "Cannot find cred variable"}
    if(!$subfolder){Write-LogEntry -LogName $logpath -LogEntryText "No Subfolder Detected"}
    else{Write-LogEntry -LogName $logpath -LogEntryText "Recover Folder is: $subfolder"}

    # Load Exchange web services DLL and set the service
    # Requires the EWS API downloaded to your local computer
    # Write-LogEntry -LogName $logpath -LogEntryText "Troubleshooting: $($creds)"
    # Write-LogEntry -LogName $logpath -LogEntryText "Troubleshooting: $($params)"
    $dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"    
    Import-Module $dllpath
    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013 
    $Global:service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion) 
    $convertedcred = New-Object System.Net.NetworkCredential($creds.UserName.ToString(),$creds.GetNetworkCredential().password.ToString())     
    $Global:service.Credentials = $convertedcred
    
    Write-LogEntry -LogName $logpath -LogEntryText "Service Credentials: $($Global:service.Credentials)"


    ## Set the URL (autodiscover can be used as well)
    $ewsURL = "https://outlook.office365.com/ews/Exchange.asmx"
    $uri= [system.URI]$ewsURL 
    Write-LogEntry -LogName $logpath -LogEntryText "Setting Service to: $uri" 
    $Global:service.Url = $uri 
    Write-LogEntry -LogName $logpath -LogEntryText "Service is set to connect to $($global:service.Url.AbsoluteUri)"
    #$Global:service.AutodiscoverUrl($AccountWithImpersonationRights ,{$true})
    ##Login to Mailbox with Impersonation
    # Full credit for the Below GetFolder and Create Folder Functions goes to David Berret
    # https://blogs.msdn.microsoft.com/emeamsgdev/2013/10/20/powershell-create-folders-in-users-mailboxes/

    function itemSearchOptions([string]$itemType){
        #combine the search into aggrated search filter
        $searchFilterCollectionItemType =  new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, $itemType)    
        $Global:searchFilterAggregated = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        if($itemType -eq "IPM.StickyNote" -or $itemType -eq "IPM.Task"){
            $Global:searchFilterAggregated.Add($searchFilterCollectionItemType)
        }
        else{
            $searchFilterStartDate = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, [System.DateTime]$startdate)
            $searchFilterEndDate = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, [System.DateTime]$enddate)      
            #specify our OR states for start and end dates
            $global:searchFilterCollectionStart = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::Or)
            $global:searchFilterCollectionStart.Add($searchFilterStartDate)
            #$global:searchFilterCollectionStart.Add($searchFilterStartDateCreated)
            $global:searchFilterCollectionEnd = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::Or)
            $global:searchFilterCollectionEnd.Add($searchFilterEndDate)
            #$global:searchFilterCollectionEnd.Add($searchFilterEndDateCreated)
            $Global:searchFilterAggregated.Add($searchFilterCollectionItemType)
            $Global:searchFilterAggregated.Add($global:searchFilterCollectionStart)
            $Global:searchFilterAggregated.Add($global:searchFilterCollectionEnd)
        }
    }
    Function GetFolder()
    {
        # Return a reference to a folder specified by path
        
        $RootFolder, $FolderPath = $args[0]
        
        $Folder = $RootFolder
        if ($FolderPath -ne '\')
        {
            $PathElements = $FolderPath -split '\\'
            For ($i=0; $i -lt $PathElements.Count; $i++)
            {
                if ($PathElements[$i])
                {
                    $View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0)
                    $View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
                            
                    $SearchFilterFolder = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i])
                    
                    $FolderResults = $Folder.FindFolders($SearchFilterFolder, $View)
                    if ($FolderResults.TotalCount -ne 1)
                    {
                        # We have either none or more than one folder returned... Either way, we can't continue
                        $Folder = $null
                        Write-LogEntry -LogName $logpath -LogEntryText ([string]::Format("Failed to find {0}", $PathElements[$i]))
                        Write-LogEntry -LogName $logpath -LogEntryText ([string]::Format("Requested folder path: {0}", $FolderPath))
                        break
                    }
                    
                    $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Global:service, $FolderResults.Folders[0].Id)
                }
            }
        }
        
        return $Folder
    }

    Function CreateFolders(){

        $FolderId = $args[0]
        $requiredFolder = $args[1]
        Write-LogEntry -LogName $logpath -LogEntryText "Binding to folder with id $FolderId" 
        $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Global:service,$FolderId)
        if (!$folder) { return }

            Write-Verbose "`tChecking for existence of $requiredFolder"
            $rf = GetFolder( $folder, $requiredFolder )
            if ( $rf )
            {
                Write-LogEntry -LogName $logpath -LogEntryText "`t$requiredFolder already exists" -ForegroundColor Green
                $global:recoveryFolder = $rf
            }
            Else
            {
                # Create the folder
                if (!$WhatIf)
                {
                    $rf = New-Object Microsoft.Exchange.WebServices.Data.Folder($Global:service)
                    $rf.DisplayName = $requiredFolder
                    $rf.Save($FolderId)
                    if ($rf.Id.UniqueId)
                    {
                        Write-LogEntry -LogName $logpath -LogEntryText "`t$requiredFolder created successfully" -ForegroundColor Green
                        $global:recoveryFolder = $rf
                    }
                    else{
                        Write-LogEntry -LogName $logpath -LogEntryText "`tEncountered a failure" -ForegroundColor Red
                        exit;
                    }
                }
                Else
                {
                    Write-LogEntry -LogName $logpath -LogEntryText "`t$requiredFolder would be created" -ForegroundColor Yellow
                }
            }
        
        }
        
    function moveItems(){
        $MailboxToImpersonate = $args[0]
        $pagelimit = $args[1]
        $subfolder = $args[2]
        $Global:service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate ); 
        Write-LogEntry -LogName $logpath -LogEntryText "Service set to impersonate for: $($Global:service.ImpersonatedUserId.Id)"
        

        #Get Recoverable Item Root Folder
        $rfRootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]:: RecoverableItemsRoot,$MailboxToImpersonate) 
        #Bind folder IDs to Objects
        $rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Global:service,$rfRootFolderID) 
        $ffResponse = $rfRootFolder.FindFolders($Global:FolderView) 
        # Grab Purges and Deletions subfolders
        $ffResponse = $ffResponse | Where-Object {$_.DisplayName -like "*Purges*" -or $_.DisplayName -like "*Deletions*"}

        foreach($option in $Global:recoverOptions.Keys){
            Write-LogEntry -LogName $logpath -LogEntryText "Recovery Option is $option"
            switch ($option) {
                recoverCalendar {
                    $global:recoveryFolder = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxToImpersonate)
                    itemSearchOptions "IPM.Appointment"
                    $destFolder = $global:recoveryFolder.FolderName
                    Write-LogEntry -LogName $logpath -LogEntryText "The Recover Folder is: $destFolder"

                }
                recoverTasks {
                    $global:recoveryFolder = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Tasks,$MailboxToImpersonate)
                    itemSearchOptions "IPM.Task"
                    $destFolder = $global:recoveryFolder.FolderName
                    Write-LogEntry -LogName $logpath -LogEntryText "The Recover Folder is: $destFolder"             
                }
                recoverNotes {
                    $global:recoveryFolder = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Notes,$MailboxToImpersonate)
                    itemSearchOptions "IPM.StickyNote"
                    $destFolder = $global:recoveryFolder.FolderName
                    Write-LogEntry -LogName $logpath -LogEntryText "The Recover Folder is: $destFolder"
                                   
                }
                RecoverEmail {
                    $inboxID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxToImpersonate) 
                    itemSearchOptions "IPM.Note"
                    if($subfolder){
                        Write-LogEntry -LogName $logpath -LogEntryText "`tTrying to create subfolder: $subfolder"
                        CreateFolders $inboxID $subfolder
                        $destFolder = $global:recoveryFolder.DisplayName
                        Write-LogEntry -LogName $logpath -LogEntryText "The Recover Folder is: $destFolder"     
                        $global:recoveryFolder = $global:recoveryFolder.Id                   
                    }                                     
                    else{
                        Write-LogEntry -LogName $logpath -LogEntryText "`tDefaulting to Inbox"                        
                        $global:recoveryFolder = $inboxID
                        $destFolder = $global:recoveryFolder.FolderName
                        Write-LogEntry -LogName $logpath -LogEntryText "The Recover Folder is: $destFolder"                        
                    }
                }
                Default {
                    $inboxID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxToImpersonate) 
                    itemSearchOptions "IPM.Note"
                    if($subfolder){
                        Write-LogEntry -LogName $logpath -LogEntryText "`tTrying to create subfolder: $subfolder"
                        CreateFolders $inboxID $subfolder
                        $destFolder = $global:recoveryFolder.DisplayName
                        Write-LogEntry -LogName $logpath -LogEntryText "The Recover Folder is: $destFolder"    
                        $global:recoveryFolder = $global:recoveryFolder.Id                    
                    }                                      
                    else{
                        $global:recoveryFolder = $inboxID
                        $destFolder = $global:recoveryFolder.FolderName
                        Write-LogEntry -LogName $logpath -LogEntryText "The Recover Folder is: $destFolder"                        
                    }
                }
            }

            
            foreach($folder in $ffResponse){
                Write-LogEntry -LogName $logpath -LogEntryText "DisplayName of Folder we are searching: $($folder.DisplayName)"            
                if($folder){
                    #Below we search items with, up to 1000 at a time
                    $items = $Global:service.FindItems($folder.Id,$Global:searchFilterAggregated,$Global:Itemsview)
                    Write-LogEntry -LogName $logpath -LogEntryText "`tThere are $($Items.TotalCount) items in $($folder.DisplayName)"
                    $movemethod = $items | Get-Member Move -ErrorAction SilentlyContinue
                    if($items -and $movemethod){
                        [int]$offset = 0
                        [int]$batch = 1
                        [decimal]$batches = (([int]$items.TotalCount)/($pagelimit))
                        $batches = [System.Math]::Ceiling($batches)
                        Write-LogEntry -LogName $logpath -LogEntryText "`tThere are $batches Batches"                
                        $moreitems = $true
                        do{             
                            $count = @($items.count)
                            $count = $count.Count
                            switch ($count) {
                                $pagelimit {$mailboxtomove = "`tMoving $count items, in Batch # $batch"}
                                Default {$mailboxtomove = "`tMoving $count, Final Batch # $batch"}
                            }

                            if($whatif){
                                Write-LogEntry -LogName $logpath -LogEntryText "Whatif: $mailboxtomove from $($folder.DisplayName) to  $destFolder" -ForegroundColor Yellow
                            }
                            else{
                                Write-LogEntry -LogName $logpath -LogEntryText "$mailboxtomove from  $($folder.DisplayName) to  $destFolder"
                                # This is our method by which we move items to the destination recovery folder"
                                $items.Move($global:recoveryFolder) | Out-Null
                            }

                            if($items.MoreAvailable -eq $false){
                                Write-LogEntry -LogName $logpath -LogEntryText "`tNo more items to move"
                                $moreitems = $false
                                $Global:Itemsview = new-object Microsoft.Exchange.WebServices.Data.ItemView([int]$pagelimit)
                                $Global:Itemsview.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
                                
                            }
                            else{
                                [int]$offset += $Global:Itemsview.PageSize
                                $Global:Itemsview.Offset = $offset
                                $batch++
                                $items = $Global:service.FindItems($folder.Id,$Global:searchFilterAggregated,$Global:Itemsview)
                                Write-LogEntry -LogName $logpath -LogEntryText "`tIncrementing Offset to $offset; Value is: $($Global:itemsview.Offset)"
                            }
                        }    
                        while($moreitems -eq $True){}
                    }
                    else{
                        Write-LogEntry -LogName $logpath -LogEntryText "`tNo Items to move, or missing Move method" -ForegroundColor Yellow
                    }
                }
                else{
                    Write-LogEntry -LogName $logpath -LogEntryText "Couldn't find the Purges folder for: $mailboxtoimpersonate" -ForegroundColor Red
                }    
            }
        }
    }
    ###############################################################################################################################
    #Below are static options for building Search Filters and View

    #setting property sets and item traversal
    #$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)   
    $Global:Itemsview = new-object Microsoft.Exchange.WebServices.Data.ItemView([int]$pagelimit)
    $Global:Itemsview.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow

    # Setting scope for EWS to process items 
    $Global:FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000); 
    $Global:FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep 
    #End Static Options
    ###############################################################################################################################
    
    moveItems $MailboxToImpersonate $pagelimit $subfolder

    if($Error.Count -gt 0){
        Write-LogEntry -LogName $logpath -LogEntryText "Run Finished with Errors Encountered"
        Write-LogEntry -LogName $logpath -LogEntryText "Last Error $($Error[0])"
        Write-Verbose "Move ran with Errors"
    }
    else{
        Write-LogEntry -LogName $logpath -LogEntryText "Move Ran Successfully"
    }
    $Global:service = $null
    $Global:searchFilterAggregated = $null
    $Global:Itemsview = $null
    $Global:FolderView = $null
}


#This is where we call all of our code, and is our masterthread 
#validate start and end date
try{
    [datetime]$startdate = Get-Date $startdate
    [datetime]$enddate = Get-Date $enddate
    if($startdate -gt $enddate){Write-Host "You cannot specify a Start Date that is after your End Date: Please Run Again" -ForegroundColor Yellow; Exit;}    
}
catch{
    Write-LogEntry -LogName $logpath -LogEntryText "Cannot parse date times you specified, try again with formats similar to MM/DD/YYYY, YYYY-MM-DD, etc" -ForegroundColor Yellow
    exit;
}
if(!$pagelimit){[int]$pagelimit = 100}
$lastrun = 0
if(!$threadlimit){[int]$threadlimit = 5}
if(!$subfolder){Write-LogEntry -LogName $logpath -LogEntryText "No Subfolder"}
[array]$temp = @()
foreach($m in $Mailboxes){
    $check = Get-RSJob -Name $m
    if($check){
        Write-LogEntry -LogName $logpath -LogEntryText "User $m has an existing job running, will skip" -ForegroundColor White
    }
    else{
        Write-LogEntry -LogName $logpath -LogEntryText "Adding $m to temp array"
        [array]$temp += $m
    }
}
[array]$Mailboxes = [array]$temp

Write-LogEntry -LogName $logpath -LogEntryText "Entering While Loop to Provision Jobs" -ForegroundColor White
$exit = $false
While($exit -eq $false){

    $jobs = Get-RSJob -State Running -ErrorAction SilentlyContinue
    #Write-LogEntry -LogName $logpath -LogEntryText "There are $($jobs.Count) running"
    if($jobs.Count -lt $threadlimit){
        [int]$finalrun = (($Mailboxes.Count) - $lastrun)
        [int]$jobstoprovision = ($threadlimit - [int]($jobs.Count))
        if($finalrun -lt $jobstoprovision){
            Write-LogEntry -LogName $logpath -LogEntryText "Remain jobs is less than thread count: Mailboxes left $finalrun"
            $jobstoprovision = $finalrun
        }
        Write-LogEntry -LogName $logpath -LogEntryText "Need to provision $jobstoprovision jobs" -ForegroundColor Yellow
        $provisioned = 1
        while($provisioned -le $jobstoprovision){
            [string]$MailboxToImpersonate = [array]$Mailboxes[$lastrun]
            $paramblock = @{
                MailboxToImpersonate = $MailboxToImpersonate
                subfolder = $subfolder
                AccountWithImpersonationRights = $AccountWithImpersonationRights
                startdate = $startdate
                enddate = $enddate
                whatif = $whatif
                pagelimit = $pagelimit
                logpath = $rootpath
            }
            if($recoverTasks){$paramblock.recoverTasks = "Yes"}
            if($recoverNotes){$paramblock.recoverNotes = "Yes"}
            if($recoverCalendar){$paramblock.recoverCalendar = "Yes"}
            if($recoverAll){$paramblock.recoverAll = "Yes"}
            Write-LogEntry -LogName $logpath -LogEntryText "Provisioning Job  $MailboxToImpersonate; LastRun: $lastrun" -ForegroundColor White
            Start-RSJob -Name $MailboxToImpersonate -ScriptBlock $scriptblock
            $provisioned++
            $lastrun++
        }
    $total = (Get-RSJob).Count
    Write-LogEntry -LogName $logpath -LogEntryText "Total Jobs Provisioned: $total; Sleeping for 5 seconds"
    Start-Sleep -Seconds 3
    }
    if(($lastrun -ge $Mailboxes.Count)){
        while(Get-RSJob -State Running){
            Write-Progress -Activity "Waiting for $((Get-RSJob -State Running).Count) jobs to finish" -PercentComplete 100 -Id 0            
            Write-LogEntry -LogName $logpath -LogEntryText "`nPausing for 30 seconds, and will check again" -ForegroundColor Yellow
            Start-Sleep -Seconds 30
        }
        Write-LogEntry -LogName $logpath -LogEntryText "All Jobs Finished" -ForegroundColor Yellow
        Get-RSJob | Export-Csv $rootpath\RSJob-State.csv -NoTypeInformation
        $exit = $True
    }
    else{
        Write-Progress -Activity "Created $lastrun out of $($mailboxes.Count) Jobs" -PercentComplete (($lastrun/($Mailboxes.Count))*100) -Id 0
    }

}

Write-LogEntry -LogName $logpath -LogEntryText "Cleaning up Jobs" -ForegroundColor Yellow
Get-RSJob -State Completed | Remove-RSJob