# uninstall.ps1
# -------------------------------------------------
# Uninstalls a transport agent on an Exchange Server 2010 development server
#
# Copyright (c) 2013/2016 Thomas Stensitzki
#
# $EXDIR needs to be adjusted to the target environment

$EXDIR='D:\Program Files\Microsoft\Exchange Server\V14'

# Stop the transport service to remove the agent
Stop-Service MSExchangeTransport

# Disable the transport agent
Write-Output -InputObject "Disabling Agent..."
Disable-TransportAgent -Identity "SFTools Modify Attachment Agent" -Confirm:$false

# Uninstall the transport agent
Write-Output 'Uninstalling Agent..'
Uninstall-TransportAgent -Identity "SFTools Modify Attachment Agent" -Confirm:$false

# Restart IIS as the W3SVC service locks the agent DLL
Write-Output 'Restarting IIS'
Restart-Service w3svc

# Remove transport agent files from the file system
Write-Output 'Deleting Files and Folders...'
Remove-Item $EXDIR\TransportRoles\Agents\MessagingModifyAttachment\* -Recurse -ErrorAction SilentlyContinue
Remove-Item $EXDIR\TransportRoles\Agents\MessagingModifyAttachment -Recurse -ErrorAction SilentlyContinue

# Start the transport service
Start-Service MSExchangeTransport

# We are finished
Write-Output 'Uninstall Complete.'
