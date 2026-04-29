$ExtensionApp = ""

New-MgApplicationExtensionProperty `
    -ApplicationId $ExtensionApp `
    -Name "StartingPokemon" `
    -DataType "String" `
    -TargetObjects @("User")


Get-MgApplicationExtensionProperty -ApplicationId $ExtensionApp |
    Format-Table Name, DataType, TargetObjects




# Build the full extension property name
# Replace this with YOUR owner app's AppId (client ID), with hyphens removed
$ownerAppId = ""
$extensionName = "extension_${ownerAppId}_StartingPokemon"

# Build the params hashtable — has to be done this way because the property
# name is dynamic (contains your specific app ID)
$params = @{
    $extensionName = "Squirtle"
}

# Apply it to the user
Update-MgUser -UserId "" -BodyParameter $params