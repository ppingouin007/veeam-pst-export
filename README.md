# Veeam Backup for Microsoft 365 - PST Export Automation

[![PowerShell Version](https://img.shields.io/badge/PowerShell-7.0%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Veeam Version](https://img.shields.io/badge/Veeam-7.0%2B-green.svg)](https://www.veeam.com/fr/backup-microsoft-office-365.html)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

> **Scripts PowerShell pour l'export automatisé des boîtes aux lettres et archives en ligne vers des fichiers PST depuis Veeam Backup for Microsoft 365**

## 📋 Description

Ce dépôt contient une collection de scripts PowerShell conçus pour automatiser l'exportation des boîtes aux lettres Microsoft 365 vers des fichiers PST via **Veeam Backup for Microsoft 365**. L'ensemble des scripts permet de :

- Exporter en masse les boîtes aux lettres principales et les archives en ligne
- Générer automatiquement la liste des utilisateurs depuis Active Directory
- Gérer les utilisateurs sans adresse email
- Produire des fichiers PST nommés de manière cohérente

## 🎯 Contexte et cas d'usage

### Pourquoi ces scripts ?

L'interface graphique de Veeam Backup for Microsoft 365 ne permet pas d'exporter plusieurs boîtes aux lettres en une seule opération. Ces scripts comblent cette lacune en automatisant le processus via PowerShell, ce qui est particulièrement utile dans les situations suivantes :

| Cas d'usage | Description |
|-------------|-------------|
| **Départ d'employés** | Exporter rapidement les données d'un utilisateur avant de libérer sa licence Office 365 |
| **Conformité et eDiscovery** | Réaliser des exports ponctuels pour des besoins légaux ou d'audit |
| **Migration de données** | Préparer des archives avant une migration vers une autre solution |
| **Gestion de projet** | Exporter les boîtes aux lettres d'une équipe spécifique pour archivage |
| **Besoins de sauvegarde** | Créer des copies hors ligne des données critiques |

### Architecture de la solution

```
┌─────────────────────────────────────────────────────────────────┐
│                        Active Directory                         │
│  ┌─────────────────────────────────────────────────────────┐    │
│  │  OU spécifique (ex: OU=Departement,DC=entreprise,DC=fr)│    │
│  │  ├── Utilisateur 1 (mail: user1@entreprise.fr)         │    │
│  │  ├── Utilisateur 2 (mail: user2@entreprise.fr)         │    │
│  │  └── Utilisateur 3 (sans mail) ← listé séparément      │    │
│  └─────────────────────────────────────────────────────────┘    │
│                              │                                   │
│                              ▼                                   │
│  ┌─────────────────────────────────────────────────────────┐    │
│  │  generate-user-list.ps1                                │    │
│  │  └── users-to-archive.csv → Pour export Veeam          │    │
│  └─────────────────────────────────────────────────────────┘    │
│                              │                                   │
│                              ▼                                   │
│  ┌─────────────────────────────────────────────────────────┐    │
│  │  Export-VeeamMailboxes.ps1                             │    │
│  │  ├── Boîte principale → user_Principale.pst           │    │
│  │  └── Archive en ligne → user_Archive.pst              │    │
│  └─────────────────────────────────────────────────────────┘    │
│                              │                                   │
│                              ▼                                   │
│  ┌─────────────────────────────────────────────────────────┐    │
│  │           Dossier d'export (ex: D:\PST_Exports)        │    │
│  │  ├── user1_Principale.pst                             │    │
│  │  ├── user1_Archive.pst                                │    │
│  │  └── user2_Principale.pst                             │    │
│  └─────────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────────┘
```

## ✨ Fonctionnalités

### Script principal : `Export-VeeamMailboxes.ps1`

- ✅ Export de la boîte aux lettres principale en `.pst`
- ✅ Export de l'archive en ligne (si existante) en `.pst`
- ✅ Recherche flexible des utilisateurs (si l'email exact n'est pas trouvé)
- ✅ Journalisation détaillée des opérations
- ✅ Gestion des erreurs et résumé des exports
- ✅ Fermeture automatique des sessions de restauration

### Script de génération de liste : `Generate-UserList.ps1`

- ✅ Récupération des utilisateurs depuis un OU Active Directory
- ✅ Séparation des utilisateurs avec/sans email
- ✅ Génération du fichier CSV

## 📦 Prérequis

### Logiciels requis

| Composant | Version | Remarque |
|-----------|---------|----------|
| **PowerShell** | 7.0 ou supérieur | Obligatoire pour Veeam v8+ |
| **Veeam Backup for Microsoft 365** | 7.0 ou supérieur | Modules PowerShell inclus |
| **Microsoft Outlook** | 64-bit (2010/2013/2016) | Requis pour l'export PST |
| **Active Directory Module** | - | Pour la génération de liste depuis AD |

### Droits requis

- Droits d'administration sur le serveur Veeam
- Droits de lecture sur Active Directory (pour `Generate-UserList.ps1`)
- Droits d'écriture sur le dossier d'export
- Droits de restauration dans Veeam Backup for Microsoft 365

### Structure de fichiers attendue

```
📁 Veeam-PST-Export/
├── 📄 Export-VeeamMailboxes.ps1
├── 📄 Generate-UserList.ps1
├── 📄 README.md
└── 📄 LICENSE
```

## 📖 Utilisation

### Étape 1 : Générer la liste des utilisateurs

```powershell
.\Generate-UserList.ps1
```
**Paramètres :**
- `-SearchBase` : Chemin de l'OU Active Directory à interroger (obligatoire)
- `-OutputPath` : Dossier de sortie des CSV (défaut : `C:\scripts`)

**Fichiers générés :**
- `users-to-archive.csv` → Pour l'export Veeam

### Étape 2 : Exécuter l'export des boîtes aux lettres

```powershell
.\Export-VeeamMailboxes.ps1 -CsvPath "C:\scripts\users-to-archive.csv" -ExportPath "D:\PST_Exports" -OrganizationName "entreprise.onmicrosoft.com"
```

**Paramètres :**
- `-CsvPath` : Chemin du fichier CSV contenant la liste des utilisateurs (format : colonne `UserPrincipalName` contenant les emails)
- `-ExportPath` : Dossier de destination des fichiers PST
- `-OrganizationName` : Nom de l'organisation dans Veeam (ex: `entreprise.onmicrosoft.com`)

**Résultat :**
```
D:\PST_Exports\
├── user1_Principale.pst
├── user1_Archive.pst
└── user2_Principale.pst
```

## 🔧 Dépannage

### Erreur : "Cannot convert 'System.Object[]'"

**Cause** : Utilisation de PowerShell 5.1 au lieu de PowerShell 7.

**Solution** : 
```powershell
# Installer PowerShell 7
winget install --id Microsoft.PowerShell --source winget

# Exécuter avec PowerShell 7
pwsh .\Export-VeeamMailboxes.ps1 ...
```

### Erreur : "Boîte principale non trouvée"

**Cause** : L'email dans le CSV ne correspond pas à celui dans Veeam.

**Solution** : Vérifier le format d'email exact utilisé dans Veeam :

```powershell
# Une fois la session démarrée
$allMailboxes | Select-Object Email, DisplayName | Format-Table -AutoSize
```

### Erreur : "Impossible de charger l'assembly"

**Cause** : Les modules Veeam ne sont pas installés pour PowerShell 7.

**Solution** : Réinstaller Veeam Backup for Microsoft 365 ou copier les modules :

```powershell
# Vérifier les modules disponibles
Get-Module -Name Veeam.* -ListAvailable

# Si manquants, réparer l'installation de Veeam
```

### L'export reste bloqué

- Vérifier que Outlook 64-bit est bien installé et configuré
- S'assurer qu'aucune session Outlook n'est en cours
- Vérifier l'espace disque disponible sur le dossier d'export

## 📝 Exemple complet

```powershell
# 1. Générer la liste depuis l'OU "Departement"
.\Generate-UserList.ps1

# 2. Vérifier le fichier généré
Get-Content "C:\scripts\users-to-archive.csv"

# 3. Lancer l'export
.\Export-VeeamMailboxes.ps1 -CsvPath "C:\scripts\users-to-archive.csv" -ExportPath "D:\Exports\PST" -OrganizationName "entreprise.onmicrosoft.com"
```

## 🤝 Contribution

Les contributions sont les bienvenues ! 

## 📄 Licence

Ce projet est sous licence MIT - voir le fichier [LICENSE](LICENSE) pour plus de détails.

## 🙏 Remerciements

- Veeam pour la mise à disposition de leurs modules PowerShell
- La communauté PowerShell pour les bonnes pratiques
- Les contributeurs du projet

---

**📌 Note importante** : Ces scripts sont fournis "en l'état", sans garantie explicite ou implicite. Testez-les dans un environnement de développement avant toute utilisation en production.

---

## 📋 Modèle de fichier CSV attendu

Votre fichier CSV doit contenir une colonne nommée `UserPrincipalName` avec les adresses email des utilisateurs :

```csv
UserPrincipalName
prenom.nom@entreprise.fr
prenom.nom@entreprise.fr
prenom.nom@entreprise.fr
```

> ⚠️ **Important** : Utilisez l'attribut `mail` depuis Active Directory, pas `UserPrincipalName`. Voir la section [Traitement des utilisateurs sans email](#-traitement-des-utilisateurs-sans-email) pour plus de détails.

---

## 📊 Tableau des scripts

| Script | Fonction | Entrée | Sortie |
|--------|----------|--------|--------|
| `Generate-UserList.ps1` | Génération de la liste des utilisateurs | OU Active Directory | CSV (avec/sans email) |
| `Export-VeeamMailboxes.ps1` | Export des boîtes aux lettres | CSV (emails) | Fichiers PST |

---

## 🔒 Sécurité

- Les scripts utilisent l'authentification Veeam existante
- Aucun mot de passe n'est stocké dans les scripts
- Les fichiers CSV ne contiennent que des adresses email

---
