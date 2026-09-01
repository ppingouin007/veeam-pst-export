<#
.SYNOPSIS
    Exporte automatiquement les boîtes aux lettres principales et les archives en ligne
    d'une liste d'utilisateurs vers des fichiers PST.
.DESCRIPTION
    Ce script utilise les modules PowerShell de Veeam Backup for Microsoft 365 pour
    exporter, pour chaque utilisateur listé dans un fichier CSV, sa boîte aux lettres
    principale et son archive en ligne (si elle existe) vers des fichiers PST distincts.
.NOTES
    Prérequis : 
    - PowerShell 7 ou plus récent
    - Modules Veeam.Archiver.PowerShell et Veeam.Exchange.PowerShell importés
    - Microsoft Outlook 64-bit (2010, 2013 ou 2016) installé
    - Le compte exécutant le script doit avoir les droits nécessaires dans Veeam
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory = $true)]
    [string]$ExportPath,
    
    [Parameter(Mandatory = $true)]
    [string]$OrganizationName
)

# --- 1. Importation des modules Veeam ---
try {
    Import-Module Veeam.Archiver.PowerShell -ErrorAction Stop
    Import-Module Veeam.Exchange.PowerShell -ErrorAction Stop
    Write-Host "Modules Veeam importés avec succès." -ForegroundColor Green
} catch {
    Write-Error "Erreur lors de l'importation des modules Veeam. Erreur : $($_.Exception.Message)"
    exit 1
}

# --- 2. Création du dossier d'export si inexistant ---
if (-not (Test-Path -Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
    Write-Host "Dossier d'export créé : $ExportPath" -ForegroundColor Yellow
}

# --- 3. Lecture du fichier CSV ---
try {
    $userList = Import-Csv -Path $CsvPath -ErrorAction Stop
    Write-Host "$($userList.Count) utilisateur(s) chargé(s) depuis le fichier CSV." -ForegroundColor Green
} catch {
    Write-Error "Impossible de lire le fichier CSV à l'emplacement '$CsvPath'. Erreur : $($_.Exception.Message)"
    exit 1
}

# --- 4. Connexion et démarrage de la session de restauration ---
try {
    Write-Host "Récupération de l'organisation '$OrganizationName'..." -ForegroundColor Cyan
    $organization = Get-VBOOrganization -Name $OrganizationName -ErrorAction Stop

    Write-Host "Démarrage d'une session de restauration Exchange pour l'organisation..." -ForegroundColor Cyan
    
    # Fermer toute session existante
    $existingSessions = Get-VBOExchangeItemRestoreSession
    if ($existingSessions) {
        Write-Host "Fermeture des sessions existantes..." -ForegroundColor Yellow
        foreach ($session in $existingSessions) {
            Stop-VBOExchangeItemRestoreSession -Session $session -ErrorAction SilentlyContinue
        }
    }
    
    # Démarrer une nouvelle session
    $session = Start-VBOExchangeItemRestoreSession -LatestState -Organization $organization -ErrorAction Stop
    Write-Host "Session de restauration démarrée avec succès." -ForegroundColor Green
    Write-Host "ID Session : $($session.Id)" -ForegroundColor Gray
    
    # Récupérer la base de données de la session
    Write-Host "Récupération de la base de données de sauvegarde..." -ForegroundColor Cyan
    $database = $session | Get-VEXDatabase -ErrorAction Stop
    
    if (-not $database) {
        Write-Error "Aucune base de données trouvée dans la session. Vérifiez que des sauvegardes sont disponibles."
        exit 1
    }
    
    Write-Host "Base de données récupérée avec succès." -ForegroundColor Green
    
    # Récupération de toutes les boîtes aux lettres
    Write-Host "Récupération de toutes les boîtes aux lettres disponibles..." -ForegroundColor Cyan
    $allMailboxes = $database | Get-VEXMailbox -ErrorAction Stop
    
    if (-not $allMailboxes) {
        Write-Error "Aucune boîte aux lettres trouvée dans la base de données."
        exit 1
    }
    
    Write-Host "Session de restauration prête. $($allMailboxes.Count) boîtes aux lettres trouvées." -ForegroundColor Green
    
} catch {
    Write-Error "Erreur lors de la configuration de la session de restauration. Erreur : $($_.Exception.Message)"
    Write-Host "Détails de l'erreur : $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}

# --- 5. Boucle d'exportation pour chaque utilisateur ---
$successCount = 0
$failCount = 0
$failedUsers = @()

foreach ($user in $userList) {
    $upn = $user.UserPrincipalName
    Write-Host "`n--- Traitement de l'utilisateur : $upn ---" -ForegroundColor Cyan
    
    # Nettoyage du nom pour le nom de fichier
    $cleanName = $upn -replace '@.*$' -replace '[_\s]+', '.'
    $principalPstPath = Join-Path -Path $ExportPath -ChildPath "$cleanName`_Principale.pst"
    $archivePstPath = Join-Path -Path $ExportPath -ChildPath "$cleanName`_Archive.pst"

    try {
        # 5.1. Récupération de la boîte aux lettres principale
        $principalMailbox = $allMailboxes | Where-Object { 
            $_.Email -eq $upn -and $_.IsArchive -eq $false 
        }
        
        if ($principalMailbox) {
            Write-Host "  Boîte principale trouvée pour $upn." -ForegroundColor Green
            Write-Host "  Export de la boîte principale vers : $principalPstPath" -ForegroundColor Yellow
            
            # Export de la boîte aux lettres principale
            $principalMailbox | Export-VEXItem -To $principalPstPath -ErrorAction Stop
            Write-Host "  Export principal terminé avec succès." -ForegroundColor Green
        } else {
            Write-Warning "  Boîte principale non trouvée pour $upn."
            $failCount++
            $failedUsers += "$upn (principale)"
            continue
        }

        # 5.2. Récupération de l'archive en ligne
        $archiveMailbox = $allMailboxes | Where-Object { 
            $_.Email -eq $upn -and $_.IsArchive -eq $true 
        }
        
        if ($archiveMailbox) {
            Write-Host "  Archive en ligne trouvée pour $upn." -ForegroundColor Green
            Write-Host "  Export de l'archive vers : $archivePstPath" -ForegroundColor Yellow
            
            $archiveMailbox | Export-VEXItem -To $archivePstPath -ErrorAction Stop
            Write-Host "  Export de l'archive terminé avec succès." -ForegroundColor Green
        } else {
            Write-Host "  Aucune archive en ligne trouvée pour $upn." -ForegroundColor Gray
        }
        
        $successCount++

    } catch {
        Write-Warning "  Erreur lors du traitement de l'utilisateur '$upn'. Erreur : $($_.Exception.Message)"
        $failCount++
        $failedUsers += "$upn (erreur: $($_.Exception.Message))"
    }
}

# --- 6. Nettoyage et fermeture de la session ---
try {
    Write-Host "`nFermeture de la session de restauration..." -ForegroundColor Cyan
    Stop-VBOExchangeItemRestoreSession -Session $session -ErrorAction SilentlyContinue
    Write-Host "Session fermée avec succès." -ForegroundColor Green
} catch {
    Write-Warning "Erreur lors de la fermeture de la session : $($_.Exception.Message)"
}

# --- 7. Résumé final ---
Write-Host "`n=========================================" -ForegroundColor Cyan
Write-Host "Résumé des exports :" -ForegroundColor Cyan
Write-Host "  - Exportations réussies : $successCount utilisateurs" -ForegroundColor Green
Write-Host "  - Exportations échouées : $failCount utilisateurs" -ForegroundColor Red

if ($failedUsers.Count -gt 0) {
    Write-Host "`nListe des utilisateurs échoués :" -ForegroundColor Yellow
    foreach ($failed in $failedUsers) {
        Write-Host "  - $failed" -ForegroundColor Yellow
    }
}

Write-Host "=========================================" -ForegroundColor Cyan
