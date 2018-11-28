New-PnPList -Title 'NewUsers' -Template GenericList -Url Lists/NewUsers
Add-PnPField -List "NewUsers" -AddToDefaultView -DisplayName "UPN" -InternalName "UPN" -Type Text
Add-PnPField -List "NewUsers" -AddToDefaultView -DisplayName "UserManager" -InternalName "UserManager" -Type User
Add-PnPField -List "NewUsers" -AddToDefaultView -DisplayName "Department" -InternalName "Department" -Type Choice -Choices "Dep1","Dep2", "Dep3"
Add-PnPField -List "NewUsers" -AddToDefaultView -DisplayName "GivenName" -InternalName "GivenName" -Type Text
Add-PnPField -List "NewUsers" -AddToDefaultView -DisplayName "SurName" -InternalName "SurName" -Type Text
Add-PnPField -List "NewUsers" -AddToDefaultView -DisplayName "Jobtitle" -InternalName "Jobtitle" -Type Choice -Choices "Jobtitle1","Jobtitle2", "Jobtitle3"
Add-PnPField -List "NewUsers" -AddToDefaultView -DisplayName "UsageLocation" -InternalName "UsageLocation" -Type Choice -Choices "DE","EN"
Add-PnPField -List "NewUsers" -AddToDefaultView -DisplayName "License" -InternalName "License" -Type Choice -Choices "E1","E3", "Stream", "PowerApps", "Flow", "PowerBI"
Add-PnPField -List "NewUsers" -AddToDefaultView -DisplayName "MailAddress" -InternalName "MailAddress" -Type Text