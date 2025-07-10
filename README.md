# Script de Connexion Exchange Online
# 🔐 Script PowerShell pour se connecter à Exchange Online via le Partner Center Microsoft 365 (multi-tenant).

Ce script PowerShell permet de se connecter facilement à Exchange Online en tant qu'administrateur délégué via le Centre Partenaire Microsoft 365. Il a été conçu pour simplifier la gestion des clients et l'exécution de commandes Exchange dans un environnement multi-tenant.

## ✨ Fonctionnalités

- Connexion automatique au Partner Center
- Filtrage et affichage des clients sous forme de tableau
- Sélection intuitive du tenant
- Connexion à Exchange Online dans une nouvelle fenêtre PowerShell
- Journalisation des actions dans un fichier `.log`

## 🛠️ Prérequis

- PowerShell 5.1 ou supérieur
- Modules :
  - `PartnerCenter`
  - `ExchangeOnlineManagement`

## 🚀 Utilisation

1. Cloner le dépôt (ou télécharger le zip):
   ```bash
   git clone https://github.com/votre-utilisateur/exchange-online-connexion-script.git
   ```
2. Personnaliser la variable au début du script $defaultUsername :
   Utiliser de préférence Notepad++ pour modifier le fichier (Encodage UTF-8-BOM)
   ```bash
   $defaultUsername = 'agent365@mondomaine.fr' # À personnaliser
   ```
3. Lancer le script :
   ```bash
   .\Connection_Exchange_Online.ps1
   ```
