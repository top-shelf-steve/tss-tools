# Frequent requesters — dot-source before running scripts that need an owner UPN
# Usage: . .\AppRegPeople.ps1
#        .\Create-NewEntraAppRegistration.ps1 -DisplayName "MyApp" -OwnerUPN $myAppPeople.Wallace
#        .\Create-NewEntraAppRegistration.ps1 -DisplayName "TeamApp" -OwnerUPN $myAppPeople.TeamA

$myAppPeople = @{
    Wallace  = "wallace@domain.com"
    # Jordan  = "jordan@domain.com"
    # Priya   = "priya@domain.com"

    TeamA = @("blue@pallet.com", "red@pallet.com")
    TeamB = @("", "")
}


