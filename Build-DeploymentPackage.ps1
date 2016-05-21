# Build-DeploymentPackage.ps1
# -------------------------------------------------------
# Script to build the deployment package for distribution

# Create Package directory
Write-Output "Creating directory"

# Create a new folder in the current location
New-Item -Type Directory -path .\DeploymentAgentFiles -ErrorAction SilentlyContinue

# Copy required files to destination folder
Write-Output "Copying files"
Copy-Item AttachmentModify\bin\release\SFTools.Messaging.AttachmentModify.dll .\DeploymentAgent -force
Copy-Item AttachmentModify\bin\release\SFTools.Messaging.AttachmentModify.pdb .\DeploymentAgent -force
Copy-Item deployment-config.xml .\DeploymentAgent\SFTools.MessageModify.Config.xml -force
Copy-Item Add-TransportAgent.ps1 .\DeploymentAgent -force
Copy-Item Remove-TransportAgent.ps1 .\DeploymentAgent -force
Copy-Item readme.txt .\DeploymentAgent -force

Write-Output "Finished."