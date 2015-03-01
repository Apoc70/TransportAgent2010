<# 
    .SYNOPSIS 
    Uninstall a transport agent on an Exchange Server 2010

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.0, 2013

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 
    More information can be found at http://www.sf-tools.net/Messaging/ExchangeServer2010/TransportAgent/tabid/146/Default.aspx

    .DESCRIPTION 
    This script uninstalls an installled transport agent on a local Exchange Server 2010. 
    The C# example code can be found at http://www.sf-tools.net/Messaging/ExchangeServer2010/TransportAgent/tabid/146/Default.aspx 
 
    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1  
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 

    .EXAMPLE 
    .\Remove-TransportAgent.ps1

#> 

# Assuming Exchange Server 2010 is installed in the default location
$EXDIR="C:\Program Files\Microsoft\Exchange Server\V14"  
  
# Stop the transport service to remove the agent  
Stop-Service MSExchangeTransport  
  
# Disable the transport agent  
Write-Output "Disabling Agent..."  
Disable-TransportAgent -Identity "SFTools Modify Attachment Agent" -Confirm:$false  
  
# Uninstall the transport agent  
Write-Output "Uninstalling Agent.."  
Uninstall-TransportAgent -Identity "SFTools Modify Attachment Agent" -Confirm:$false  
  
# Restart IIS as the W3SVC service locks the agent DLL  
Write-Output "Restarting IIS"  
Restart-Service w3svc  
  
# Remove transport agent files from the file system  
Write-Output "Deleting Files and Folders..."  
Remove-Item $EXDIR\TransportRoles\Agents\MessagingModifyAttachment\* -Recurse -ErrorAction SilentlyContinue  
Remove-Item $EXDIR\TransportRoles\Agents\MessagingModifyAttachment -Recurse -ErrorAction SilentlyContinue  
  
# Start the transport service  
Start-Service MSExchangeTransport  
  
# We are finished  
Write-Output "Uninstall Complete"  