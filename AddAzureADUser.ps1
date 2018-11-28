# POST method: $req
$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$name = $requestBody.name

# GET method: each querystring parameter is its own variable
if ($req_query_name) 
{
    $name = $req_query_name 
}
if ($req_query_ItemID) 
{
    $itemID = $req_query_ItemID 
}
if ($req_query_URL) 
{
    $url = $req_query_Url
}
if ($req_query_ListTitle) 
{
    $listTitle = $req_query_ListTitle
}

Out-File -Encoding Ascii -FilePath $res -inputObject "Hello $name"


# Create Context for PowerShell Modules and User Credentials (connection to O365, O365 Admin)
$FunctionName = 'AddAzureADUser'

# Define Modules
$PnPModuleName = 'SharePointPnPPowerShellOnline'
$PnPVersion = '2.20.1711.0'
$AzureADModuleName = 'AzureAD'
$AzureADVersion = '2.0.0.131'
$MSOLModuleName ='MSOnline'
$MSOLVersion ='1.1.166.0'

$username = $Env:user
$pw = $Env:password

# Import PS modules
$AzureADModulePath = "D:\home\site\wwwroot\$FunctionName\bin\$AzureADModuleName\$AzureADVersion\$AzureADModuleName.psd1"
$MSOLModulePath = "D:\home\site\wwwroot\$FunctionName\bin\$MSOLModuleName\$MSOLVersion\$MSOLModuleName.psd1"
$PnPModulePath = "D:\home\site\wwwroot\$FunctionName\bin\$PnPModuleName\$PnPVersion\$PnPModuleName.psd1"
$res = "D:\home\site\wwwroot\$FunctionName\bin"
 
Import-Module $AzureADModulePath
Import-Module $PnPModulePath
Import-Module $MSOLModulePath
 
# Build Credentials
$keypath = "D:\home\site\wwwroot\$FunctionName\bin\keys\PassEncryptKey.key"
$pwfile = @(Get-Content $keypath)[0]
$secpassword = $pw | ConvertTo-SecureString -Key $pwfile 
$credentials= New-Object System.Management.Automation.PSCredential ($username, $secpassword)

# Your Tenant ID
$tenant = "TENANT ID"

# Connect to MSOL
Connect-MsolService -Credential $credentials

# Connect to SharePoint Online Service
Connect-PnPOnline -Url $url -Credentials $credentials
$item = Get-PNPListItem -List Lists/$listTitle -Id $itemId

# Connect to Azure AD
Connect-AzureAD -TenantId $tenant -Credential $credentials # Connect-AzureAD clears the password

# Create User
if($item.FieldValues.UPN)
{
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = $item.FieldValues.Password
    $PasswordProfile.ForceChangePasswordNextLogin = $true

    if($item.FieldValues.MailAddress)
    {
        $split = $item.FieldValues.MailAddress.Split("@")
        $MailNickName = $split[0]
    }

    New-AzureADUser -UserPrincipalName $item.FieldValues.UPN -DisplayName $item.FieldValues.Title -PasswordProfile $PasswordProfile -MailNickName $MailNickName -AccountEnabled $true
    Start-Sleep -Seconds "900"

    if($item.FieldValues.UserManager)
    {
        # Get Manager Object Id from Azure AD
        $itemManagerEmail = $item.FieldValues.UserManager.Email
        #$userManagerAzure = Get-AzureADUser -ObjectId $item.FieldValues.UserManager.Email
        $userManagerAzure = Get-AzureADUser -Filter "OtherMails eq '$itemManagerEmail'"
        # Set Manager in Azure AD
        Set-AzureADUserManager -ObjectId $item.FieldValues.UPN -RefObjectId $userManagerAzure.ObjectId
        # Set Manager in SharePoint Online
        $userManagerSPO = Get-PnPUserProfileProperty -Account $userManagerAzure.Mail
        Set-PnPUserProfileProperty -Account $item.FieldValues.UPN -Property "Manager" -Value $userManagerSPO.AccountName
    }
    if($item.FieldValues.MailAddress)
    {
        Set-AzureADUser -ObjectId $item.FieldValues.UPN -OtherMails $item.FieldValues.MailAddress
        Set-PnPUserProfileProperty -Account $item.FieldValues.UPN -Property "WorkEmail" -Value $item.FieldValues.MailAddress
    }
}else 
{
    Set-PnPListItem -List $listTitle -Identity $itemID -Values @{"Log" = "UPN not set correctly"}
}


if($item.FieldValues.Department)
{
    Set-AzureADUser -ObjectId $item.FieldValues.UPN -Department $item.FieldValues.Department.Label
    Set-PnPUserProfileProperty -Account $item.FieldValues.UPN -Property "Department" -Value $item.FieldValues.Department.Label
}
if($item.FieldValues.GivenName)
{
    Set-AzureADUser -ObjectId $item.FieldValues.UPN -GivenName $item.FieldValues.GivenName
}
if($item.FieldValues.SurName)
{
    Set-AzureADUser -ObjectId $item.FieldValues.UPN -Surname $item.FieldValues.SurName
}
if($item.FieldValues.Jobtitle)
{
    Set-AzureADUser -ObjectId $item.FieldValues.UPN -JobTitle $item.FieldValues.Jobtitle.Label
    Set-PnPUserProfileProperty -Account $item.FieldValues.UPN -Property "SPS-JobTitle" -Value $item.FieldValues.Jobtitle.Label
}
if($item.FieldValues.UsageLocation)
{
    Set-AzureADUser -ObjectId $item.FieldValues.UPN -UsageLocation $item.FieldValues.UsageLocation
}

if($item.FieldValues.License)
{
    # STANDARDPACK = E1
    if($item.FieldValues.License -eq "E1")
    {
        Set-MsolUserLicense -UserPrincipalName $item.FieldValues.UPN -AddLicenses "TENANTNAME:STANDARDPACK"
    }
    # ENTERPRISEPACK = E3
    if($item.FieldValues.License -eq "E3")
    {
        Set-MsolUserLicense -UserPrincipalName $item.FieldValues.UPN -AddLicenses "TENANTNAME:ENTERPRISEPACK"
    }
    # STREAM = Stream
    if($item.FieldValues.License -eq "Stream")
    {
        Set-MsolUserLicense -UserPrincipalName $item.FieldValues.UPN -AddLicenses "TENANTNAME:STREAM"
    }
    # POWERAPPS_INDIVIDUAL_USER = PowerApps
    if($item.FieldValues.License -eq "PowerApps")
    {
        Set-MsolUserLicense -UserPrincipalName $item.FieldValues.UPN -AddLicenses "TENANTNAME:POWERAPPS_INDIVIDUAL_USER"
    }
    # FLOW_FREE = Flow & Logic
    if($item.FieldValues.License -eq "Flow")
    {
        Set-MsolUserLicense -UserPrincipalName $item.FieldValues.UPN -AddLicenses "TENANTNAME:FLOW_FREE"
    }
    # POWER_BI_STANDARD = Power BI
    if($item.FieldValues.License -eq "PowerBI")
    {
        Set-MsolUserLicense -UserPrincipalName $item.FieldValues.UPN -AddLicenses "TENANTNAME:POWER_BI_STANDARD"
    }
}