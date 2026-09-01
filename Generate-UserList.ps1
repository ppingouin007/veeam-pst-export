# Version simplifiée pour une exécution rapide
$SearchBase = "OU=Archive,OU=Users,OU=Organization,DC=domain,DC=local"
$OutputPath = "C:\scripts"

# Récupération
$users = Get-ADUser -Filter * -SearchBase $SearchBase -Properties mail, displayName

# Séparation
$withEmail = $users | Where-Object { $_.mail }
$withoutEmail = $users | Where-Object { -not $_.mail }

# Export pour Veeam
$withEmail | Select-Object @{N="UserPrincipalName";E={$_.mail}} | 
Export-Csv "$OutputPath\users-to-archive.csv" -NoTypeInformation

# Affichage des utilisateurs sans email
if ($withoutEmail) {
    Write-Host "`nATTENTION : $($withoutEmail.Count) utilisateur(s) sans email :" -ForegroundColor Yellow
    $withoutEmail | Select-Object SamAccountName, Name | Format-Table -AutoSize
}