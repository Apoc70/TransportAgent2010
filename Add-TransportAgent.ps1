<# 
    .SYNOPSIS 
    Install a transport agent on an Exchange Server 2010

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.0, 2013

    Please send ideas, comments and suggestions to support@granikos.eu 

    .DESCRIPTION 
    This script installs a transport agent on a local Exchange Server 2010. 
    The C# example code can be found at http://www.sf-tools.net/Messaging/ExchangeServer2010/TransportAgent/tabid/146/Default.aspx 
 
    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1  
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 

    .EXAMPLE 
    .\Add-TransportAgent.ps1

#> 
# Assuming Exchange Server 2010 is installed in the default location 
$EXDIR="C:\Program Files\Microsoft\Exchange Server\V14"  
      
# Stop the ExchangeTransport Service to add the new transport agent  
Stop-Service MSExchangeTransport  
      
# Create the required transport agent directory  
Write-Output "Creating directories"  
New-Item -Type Directory -path $EXDIR\TransportRoles\Agents\MessagingModifyAttachment -ErrorAction SilentlyContinue  
      
# Copy the transport agent library and the config file to the target directory  
# Source files mus be in the same folder as this script  
Write-Output "Copying files"  
Copy-Item SFTools.Messaging.AttachmentModify.dll $EXDIR\TransportRoles\Agents\MessagingModifyAttachment -force  
Copy-Item SFTools.Messaging.AttachmentModify.pdb $EXDIR\TransportRoles\Agents\MessagingModifyAttachment -force  
Copy-Item config.xml $EXDIR\TransportRoles\Agents\MessagingModifyAttachment -force  
      
# Register transport agent with Exchange   
Write-Output "Registering agent"  
Install-TransportAgent -Name "SFTools Modify Attachment Agent" -AssemblyPath $EXDIR\TransportRoles\Agents\MessagingModifyAttachment\SFTools.Messaging.AttachmentModify.dll -TransportAgentFactory SFTools.Messaging.AttachmentModify.MessageModifierFactory  
      
# Enable the transport agent  
Write-Output "Enabling agent"  
Enable-TransportAgent -Identity "SFTools Modify Attachment Agent"  
Get-TransportAgent -Identity "SFTools Modify Attachment Agent"  
      
# Start transport agent service  
Start-Service MSExchangeTransport  
      
# We are finished  
Write-Output "Install Complete. Please exit the Exchange Management Shell."  