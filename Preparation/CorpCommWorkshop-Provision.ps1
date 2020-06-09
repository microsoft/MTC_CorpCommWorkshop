# PREREQUISITES
#Install-Module -Name MSOnline
#Install-Module -Name MicrosoftTeams
#Install-Module -Name Microsoft.Online.SharePoint.PowerShell
#Install-Module SharepointPNPPowershellOnline

#TODO: make these parameters instead?
$orgName = Read-Host "Enter your Organization (i.e. ORGANIZATION.onmicrosoft.com)"
$usersFile = Read-Host "Enter path to users csv file"
$userCredential = Get-Credential

Write-Host "Connecting to AAD"
Connect-MsolService -Credential $userCredential

Write-Host "Connecting to Microsoft Teams"
Connect-MicrosoftTeams -Credential $userCredential

Write-Host "Connecting to SharePoint Online"
Connect-SPOService -Url https://$orgName-admin.sharepoint.com -Credential $userCredential
Connect-PnPOnline -Url "https://$orgName.sharepoint.com/sites/Contoso" -Credential $userCredential


#Write-Host "Disable SSPR to prevent additional factor prompts"
#TODO: NOT THE RIGHT CALL: Set-MsolCompanySettings -SelfServePasswordResetEnabled $False 


Write-Host "Creating Organization Assets libraries in Contoso site"
New-PnPList -Title "Brand Logos" -Url "BrandLogos" -Template DocumentLibrary -OnQuickLaunch
#New-PnPList -Title "Stock Photos" -Url "StockPhotos" -Template DocumentLibrary -OnQuickLaunch


#Write-Host "Upload images to Organization Asset libraries"
#$fso1 = New-Object -com Scripting.FileSystemObject
#$folder1 = $fso1.GetFolder("$pwd\StockPhotos")
#foreach($f in $folder1.files) {
#    Add-PnPFile -Path $f.path -Folder "StockPhotos"
#}

$fso2 = New-Object -com Scripting.FileSystemObject
$folder2 = $fso2.GetFolder("$pwd\BrandLogos")
foreach($f in $folder2.files) {
    Add-PnPFile -Path $f.path -Folder "BrandLogos"
}


Write-Host "Setting Organization Assets library"
Add-SPOOrgAssetsLibrary -LibraryURL "https://$orgName.sharepoint.com/sites/Contoso/BrandLogos" -ThumbnailURL "https://$orgName.sharepoint.com/sites/Contoso/BrandLogos/Contoso_logo.jpg"
#Add-SPOOrgAssetsLibrary -LibraryURL "https://$orgName.sharepoint.com/sites/Contoso/StockPhotos" -ThumbnailURL "https://$orgName.sharepoint.com/sites/Contoso/StockPhotos/camera.jpg"


Write-Host "Creating 5 cohort teams, 5 hub sites, uploading notebooks"
For ($i = 1; $i -lt 6; $i++)
{
    # create team for cohort
    New-Team -DisplayName "Cohort $i" -MailNickName "Cohort$i" -Visibility "private"

    #TODO: Upload cohort notebooks

    #TODO: Add tab to notebook in team General channel

    # create intranet (communications site) for cohort
    New-PnpSite -Title "Intranet - Cohort $i" -Type CommunicationSite -Url "https://$orgName.sharepoint.com/sites/Intranet-Cohort$i"
    Register-SPOHubSite -Site "https://$orgName.sharepoint.com/sites/Intranet-Cohort$i" -Principals $null
}

Write-Host "Looping through csv, adding user as owner of team and hub site"
Import-Csv $usersFile | Foreach-Object{
    $upn = $_."UserPrincipalName"
    $pwd = "pass@word1"
    $cohort = $_."Cohort"
    
    Write-Host "$upn (Cohort $cohort)"
    
    # change password
    Set-MsolUserPassword -UserPrincipalName $upn -NewPassword $pwd -ForceChangePassword $False

    # remove Global Admin role (other users beside MOD Admin are given Global admin when provisioning tenant)
    Remove-MsolRoleMember -RoleName "Company Administrator" -RoleMemberType User -RoleMemberEmailAddress $upn

    # add to team as Owner
    $team = Get-Team -DisplayName "Cohort $cohort"
    Add-TeamUser -GroupId $team.GroupId -User $upn -Role "Owner"

    # add to communications site as owner
    Add-SPOUser -Site "https://$orgName.sharepoint.com/sites/Intranet-Cohort$cohort" -LoginName $upn -Group "Intranet - Cohort $cohort Owners"
}
