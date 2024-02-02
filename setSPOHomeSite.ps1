#################################################################
# Set SharePoint Online Homesite
# Martin Löffler 
# 02.02.2022
# ADMIN ROLES: SharePoint Admin
# used sources: https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/set-spohomesite?view=sharepoint-ps
# Version: 1.0
#################################################################

# check PowerShell Version / version 5.x needed (not newer) 
$PSVersionTable

# check if this module is installed and which version (tested with version 16.0.24510.12000)
Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell 

# uninstall the module (if needed)
Uninstall-Module -Name Microsoft.Online.SharePoint.PowerShell

# install the module (if needed, run Powershell with Admin rights) 
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber

# list all commands of the SharePoint Online PowerShell Module (just for information)
Get-Command -Module Microsoft.Online.SharePoint.PowerShell


# connect to SharePoint Online of the tenant (interactive, SP Admin rights needed)
Connect-SPOService
# tenant URL: https://<tenant>-admin.sharepoint.com

# check the current homesite and VivaConnectionsDefaultStart setting (command only lists 1, when working with multiple Homesites)
Get-SPOHomeSite

# set homesite and VivaConnectionsDefaultStart setting
Set-SPOHomeSite -HomeSiteUrl "https://contoso.sharepoint.com/sites/homesite" -VivaConnectionsDefaultStart:$true

# Sets the home site to the provided site collection url and keeps the Viva Connections landing experience to the SharePoint home site.
# -HomeSiteUrl / Sets the URL of the site collection to be the home site.
# -VivaConnectionsDefaultStart / When set to $true, the VivaConnectionsDefaultStart parameter will keep the Viva Connections landing experience to the SharePoint home site for the desktop experience. If set to $false, the Viva Connections home experience will be used. This command doesn't impact the mobile experience.

# example 
Set-SPOHomeSite -HomeSiteUrl "https://martinloeffler.sharepoint.com/sites/home" -VivaConnectionsDefaultStart:$false
