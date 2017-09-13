# Move-FromPurges.ps1
# Last Edited: 9/13/2017
# By Brad
#
# DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

param(
	[Parameter(Position=0,Mandatory=$True,HelpMessage="Specifies the mailbox(es) to be accessed")]
	[ValidateNotNullOrEmpty()]
    [array]$Mailboxes,
	[Parameter(Position=1,Mandatory=$false,HelpMessage="Specifies the subfolder under Inbox that you want the messages moved to")]
	[ValidateNotNullOrEmpty()]
    [string]$subfolder,
	[Parameter(Position=2,Mandatory=$True,HelpMessage="Specify the UPN of the account that has impersonation rights to the mailboxes you've specified")]
	[ValidateNotNullOrEmpty()]
    [string]$AccountWithImpersonationRights, 
    [Parameter(Position=3,Mandatory=$True,HelpMessage="Specify Start Date Format YYYY-MM-DD")]
    [ValidateNotNullOrEmpty()]
    [datetime]$startdate,
    [Parameter(Position=4,Mandatory=$True,HelpMessage="Specify End Date YYYY-MM-DD")]
    [ValidateNotNullOrEmpty()]
    [datetime]$enddate,
    [Parameter(Position=5,Mandatory=$false,HelpMessage="Use -Whatif to show what the script would do")]
    [ValidateNotNullOrEmpty()]
    [switch]$whatif   
    )

# Load Exchange web services DLL and set the service
# Requires the EWS API downloaded to your local computer
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013 
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion) 

#Credential Management
$psCred = Get-Credential -UserName $AccountWithImpersonationRights -Message "Enter Impersonation Account Credentials"
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString()) 
$service.Credentials = $creds 
#clear variables
$psCred = $null
$creds = $null


## Set the URL (autodiscover can be used as well)
$ewsURL = "https://outlook.office365.com/ews/Exchange.asmx"
$uri= [system.URI]$ewsURL 
Write-Host "Connecting to: $uri" 
$service.Url = $uri 
$global:recoveryFolder = $null
#$service.AutodiscoverUrl($AccountWithImpersonationRights ,{$true})
##Login to Mailbox with Impersonation

# Full credit for the Below GetFolder and Create Folder Functions goes to David Berret
# https://blogs.msdn.microsoft.com/emeamsgdev/2013/10/20/powershell-create-folders-in-users-mailboxes/
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
						
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i])
				
				$FolderResults = $Folder.FindFolders($SearchFilter, $View)
				if ($FolderResults.TotalCount -ne 1)
				{
					# We have either none or more than one folder returned... Either way, we can't continue
					$Folder = $null
					Write-Verbose ([string]::Format("Failed to find {0}", $PathElements[$i]))
					Write-Verbose ([string]::Format("Requested folder path: {0}", $FolderPath))
					break
				}
				
				$Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $FolderResults.Folders[0].Id)
			}
		}
	}
	
	return $Folder
}

Function CreateFolders(){

    $FolderId = $args[0]
    $requiredFolder = $args[1]
    Write-Verbose "Binding to folder with id $FolderId"
    $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderId)
    if (!$folder) { return }

	    Write-Verbose "`tChecking for existence of $requiredFolder"
	    $rf = GetFolder( $folder, $requiredFolder )
	    if ( $rf )
	    {
		    Write-Host "$requiredFolder already exists" -ForegroundColor Green
            $global:recoveryFolder = $rf
	    }
	    Else
	    {
		    # Create the folder
		    if (!$WhatIf)
		    {
			    $rf = New-Object Microsoft.Exchange.WebServices.Data.Folder($service)
			    $rf.DisplayName = $requiredFolder
			    $rf.Save($FolderId)
			    if ($rf.Id.UniqueId)
			    {
				    Write-Host "`t$requiredFolder created successfully" -ForegroundColor Green
                    $global:recoveryFolder = $rf
			    }
                else{
                    Write-Host "`tEncountered a failure" -ForegroundColor Red
                    exit;
                }
		    }
		    Else
		    {
			    Write-Host "`t$requiredFolder would be created" -ForegroundColor Yellow
		    }
	    }
	
    }


foreach($MailboxToImpersonate in $Mailboxes){
    ##Define the SMTP Address of the mailbox to impersonate
    Write-Host 'Using ' $AccountWithImpersonationRights ' to Impersonate ' $MailboxToImpersonate
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate ); 
    #Get Folder IDs
    $rfRootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]:: RecoverableItemsRoot,$MailboxToImpersonate) 
    
    #Bind folder IDs to Objects
    $rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$rfRootFolderID) 

    #Specify Search Filters: Specify Date and Message Class Type
    $searchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, "IPM.Note")

    $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)   
    $Itemsview = new-object Microsoft.Exchange.WebServices.Data.ItemView(10000)

    # Setting scope for EWS to process items 
    $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000); 
    $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep 
    # Finding Sub folders in Recoverable items root of Archive mailbox (Deleted Items, Purges, Versions) 
    $ffResponse = $rfRootFolder.FindFolders($fvFolderView) 

    if($subfolder){
        Write-Host "`tTry to create subfolder: $subfolder"
        $inboxID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]:: Inbox,$MailboxToImpersonate) 
        CreateFolders $inboxID $subfolder
    }

    # traverse folders, move items
    foreach ($folder in $ffResponse.Folders) {
        #$Folder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, $true)
        if($folder.DisplayName -like "Purges"){
            $purgeFolder = $folder
            $items = $Service.FindItems($purgeFolder.Id,$searchFilter,$Itemsview)
            Write-Host "`tThere are $($Items.TotalCount) items in $($purgeFolder.DisplayName)"
            $movemethod = $items | Get-Member Move -ErrorAction SilentlyContinue
            if($items -and $movemethod){

                if($subfolder){
                    Write-host "`tThe Recovery Subfolder is: $($global:recoveryFolder.DisplayName)"
                    $items.Move($global:recoveryFolder.Id) | Out-Null
                    }
                else{
                    $items.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox) | Out-Null
                    }
                }
            else{
                Write-Host "`tNo Items to move, or missing Move method" -ForegroundColor Yellow
                }
            }
        } 

    }