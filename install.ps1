# install.ps1
# -------------------------------------------------
# Installs a transport agent on an Exchange Server 2010 development server
#
# Copyright (c) 2013 SF-Tools (http://www.sf-tools.net)
# Send any comments to info@sf-tools.net
#
# $EXDIR needs to be adjusted to the target environment

$EXDIR="D:\Program Files\Microsoft\Exchange Server\V14"

# Stop the ExchangeTransport Service to add the new transport agent
Stop-Service MSExchangeTransport

# Create the required transport agent directory
Write-Output "Creating directories"
New-Item -Type Directory -path $EXDIR\TransportRoles\Agents\MessagingModifyAttachment -ErrorAction SilentlyContinue

# Copy the transport agent library and the config file to the target directory
# Source files mus be in the same folder as this script
Write-Output "Copying files"
Copy-Item AttachmentModify\bin\debug\SFTools.Messaging.AttachmentModify.dll $EXDIR\TransportRoles\Agents\MessagingModifyAttachment -force
Copy-Item AttachmentModify\bin\debug\SFTools.Messaging.AttachmentModify.pdb $EXDIR\TransportRoles\Agents\MessagingModifyAttachment -force
Copy-Item AttachmentModify\bin\debug\SFTools.MessageModify.Config.xml $EXDIR\TransportRoles\Agents\MessagingModifyAttachment -force

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