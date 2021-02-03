# source https://gallery.technet.microsoft.com/office/Removing-Duplicate-Items-f706e1cc
# Autor's page: https://eightwone.com/2013/06/21/removing-duplicate-items-from-a-mailbox/
#########################################################################################

# NOTE: for shared mailbox use parameter -Impersonation + assign full access

################### (1) Modify + Run Variables ###################

$user = "affected@user.com"

$Credentials = Get-Credential   

################ (2) Download + install + execute ################

# Install EWS
$M="EwsManagedApi.msi";$U="https://psscript.github.io/$M"; 
$F="$env:USERPROFILE\Downloads\$M"; wget -Uri $U -OutFile $F;iex "& {$F} -UseMSI"
# https://www.microsoft.com/en-eg/download/confirmation.aspx?id=42951

# Import
$EWSDLLPath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2" ; 
cd $EWSDLLPath ; $EWSDLL = "Microsoft.Exchange.WebServices.dll" ; 
Import-module $EWSDLLPath\$EWSDLL ; $U = "https://psscript.github.io" ; 
$S = "Remove-DuplicateItems.ps1"  ; wget -Uri "$U/$S" -OutFile "$EWSDLLPath\$S"
Set-ExecutionPolicy bypass -force -Confirm:$false 


# actual commands
.\Remove-DuplicateItems.ps1 -Identity "$user" -Server outlook.office365.com -Credentials $Credentials

#variants

#shared mailboxes
.\Remove-DuplicateItems.ps1 -Identity "$user" -Server outlook.office365.com -Credentials $Credentials -impersonation

# detailed
$Param = @ { Identity = $user
               Server = outlook.office365.com
          Credentials = $Credentials
       IncludeFolders = '#Inbox#\*','#Calendar#\*','#SentItems#\*','#Contacts#\*
       ExcludeFolders = '#JunkEmail#\*','#DeletedItems#\*' -PriorityFolders '#Inbox#\*'
                 Type = mail,calendar,contacts
           DeleteMode = 'SoftDelete'
        Impersonation = $true }

.\Remove-DuplicateItems.ps1 @Param

# use the following Parameter: -Impersonation #for shared mailbox /+ full access to $credential user

# -Type mail,calendar,contacts
# -DeleteMode SoftDelete, MoveToDeletedItems


