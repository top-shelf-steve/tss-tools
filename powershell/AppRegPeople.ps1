# Frequent requesters — dot-source before running scripts that need an owner UPN
# Usage: . .\People.ps1
#        .\Create-NewEntraAppRegistration.ps1 -DisplayName "MyApp" -OwnerUPN $People.Wallace

$myAppPeople = @{
    Wallace  = "wallace@domain.com"
    # Jordan  = "jordan@domain.com"
    # Priya   = "priya@domain.com"
}
