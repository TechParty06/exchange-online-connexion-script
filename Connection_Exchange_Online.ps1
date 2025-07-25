# =============================
# Powershell Exchange Online 
# Connexion  365 MultiTenant
#            ..::: Fab :::..
# =============================

$defaultUsername = 'agent365@mondomaine.fr' # A PERSONNALISER

# Encodage utilisé.
#chcp 65001
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(65001)


# Vide le contenu de la fenêtre Powershell
Clear-Host

# Fonction des logs du script
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Out-File -Append -FilePath ".\exchange_script.log"
}

# Vérification et installation des modules nécessaires
$modules = @(
    @{ Name = 'PartnerCenter'; MinimumVersion = '4.0.0' },
    @{ Name = 'ExchangeOnlineManagement'; MinimumVersion = '3.0.0' }
)
$allModulesPresent = $true
foreach ($mod in $modules) {
    if (-not (Get-Module -ListAvailable -Name $mod.Name)) {
        $allModulesPresent = $false
    }
}
if ($allModulesPresent) {
    Write-Host "Tous les modules necessaires sont deja installes. L'execution va continuer..." -ForegroundColor Green
    Start-Sleep -Seconds 2
} else {
    foreach ($mod in $modules) {
        if (-not (Get-Module -ListAvailable -Name $mod.Name)) {
            Write-Host "Module $($mod.Name) non trouve. Installation en cours..." -ForegroundColor Yellow
            try {
                if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                    Write-Host "Elevation des privileges pour installer le module $($mod.Name)..." -ForegroundColor Cyan
                    $command = "Install-Module -Name $($mod.Name) -MinimumVersion $($mod.MinimumVersion) -Force -Scope AllUsers -AllowClobber"
                    Start-Process powershell -Verb runAs -ArgumentList "-NoProfile -Command $command"
                    Write-Host "Veuillez relancer le script apres l'installation du module." -ForegroundColor Red
                    exit
                } else {
                    Install-Module -Name $($mod.Name) -MinimumVersion $($mod.MinimumVersion) -Force -Scope AllUsers -AllowClobber
                    Write-Host "Module $($mod.Name) installe." -ForegroundColor Green
                }
            } catch {
                Write-Host "Erreur lors de l'installation du module $($mod.Name) : $($_.Exception.Message)" -ForegroundColor Red
                exit
            }
        }
    }
}

# -----------------------------------------------------
# Configuration : UPN administrateur delegue
# -----------------------------------------------------
$UsernameInput = Read-Host -Prompt "Entrez l'UPN de l'admin delegue ou appuyez sur Entree pour utiliser le compte par defaut ($defaultUsername)"
$Username = if ([string]::IsNullOrWhiteSpace($UsernameInput)) { $defaultUsername } else { $UsernameInput }

# -----------------------------------------------------
# Connexion au Partner Center
# -----------------------------------------------------
Write-Host ""
Write-Host "Connexion au Centre Partenaire Microsoft 365..." -ForegroundColor Cyan
Write-Host ""
try {
    Connect-PartnerCenter
    Write-Host ""
    Write-Host "Connecte au Centre Partenaire !" -ForegroundColor Green
    Write-Host ""
    Write-Log "Connexion reussie au Partner Center avec $Username"
} catch {
    Write-Host "Echec de la connexion : $($_.Exception.Message)" -ForegroundColor Red
    Write-Log "Erreur de connexion au Partner Center : $($_.Exception.Message)" "ERROR"
    exit
}

function Get-ClientList {
    Write-Host ""
    Write-Host "Recuperation de la liste des clients..." -ForegroundColor Cyan
    Write-Host ""
    $clients = Get-PartnerCustomer | Sort-Object Name

    $search = Read-Host -Prompt "Filtrer les clients par nom, domaine ou ID (laisser vide pour tout afficher)"
    if (-not [string]::IsNullOrWhiteSpace($search)) {
        $clients = $clients | Where-Object {
            $_.Name -like "*$search*" -or
            $_.Domain -like "*$search*" -or
            $_.TenantID -like "*$search*"
        }
    }

    $clientObj = @()
    $i = 0
    foreach ($client in $clients) {
        $clientObj += [PSCustomObject]@{
            Number      = $i
            ClientName  = $client.Name
            TenantID    = $client.TenantID
            DomainName  = $client.Domain
        }
        $i++
    }

    return ,$clientObj
}

$clientObj = Get-ClientList
Write-Host ""
Write-Host "Nombre de clients trouves : $($clientObj.Count)" -ForegroundColor Magenta

# -----------------------------------------------------
# Boucle principale
# -----------------------------------------------------
while ($true) {
    if ($clientObj.Count -eq 0) {
        Write-Host "Aucun client trouve avec ce filtre. Essayez un autre mot-cle ou tapez 'A'." -ForegroundColor Red
        $clientObj = Get-ClientList
        continue
    }

    Write-Host ""
    Write-Host "Liste des tenants disponibles :" -ForegroundColor Yellow
    Write-Host ""
    foreach ($client in $clientObj) {
        $num    = $client.Number.ToString().PadLeft(3)
        $name   = $client.ClientName
        $domain = $client.DomainName
        Write-Host "$num : $name" -ForegroundColor White
        Write-Host "      -> $domain" -ForegroundColor DarkGray
    }

    $clientChoice = Read-Host -Prompt "Selectionner le numero du tenant, taper 'A' pour actualiser, ou 'Q' pour quitter"

    switch ($clientChoice.ToLower()) {
        'q' {
            Write-Host ""
            Write-Host "Merci d'avoir utilise ce script. A bientot !" -ForegroundColor Cyan
            Write-Log "Script termine par l'utilisateur"
            exit
        }
        'a' {
            $clientObj = Get-ClientList
            Write-Host ""
            Write-Host "Nombre de clients trouves : $($clientObj.Count)" -ForegroundColor Magenta
            continue
        }
        default {
            if ($clientChoice -match '^\d+$') {
                $index = [int]$clientChoice
                if ($index -ge 0 -and $index -lt $clientObj.Count) {
                    $selectedClient = $clientObj[$index]
                    $DomainNameId = $selectedClient.DomainName
                    $ClientName = $selectedClient.ClientName
                    $windowTitle = "Exch_Online-$ClientName"

                    Write-Host ""
                    Write-Host "Connexion a '$ClientName' ($DomainNameId)..." -ForegroundColor Cyan
                    Write-Log "Connexion a $ClientName ($DomainNameId)"

                    $escapedCommand = @"
`$Host.UI.RawUI.WindowTitle = '$windowTitle' ;

# Vide le contenu de la fenêtre Powershell
`Clear-Host ;

try {
    Connect-ExchangeOnline -UserPrincipalName '$Username' -DelegatedOrganization '$DomainNameId' ;
    if (`$?) {
        Write-Host 'Connecte à $windowTitle' -ForegroundColor Green ;
        Write-Host ' ' ;
        Write-Host 'Ci-dessous les boites aux lettres trouvees :' -ForegroundColor White ;
        Write-Host ' ' ;
        Get-Mailbox -ResultSize Unlimited | Format-Table DisplayName,PrimarySmtpAddress ;
        Write-Host 'Tapez votre commande Exchange Online souhaitee' ;
        Write-Host ' ' ;
    } else {
        Write-Host 'La connexion a échoue pour $DomainNameId' -ForegroundColor Red ;
    }
} catch {
    Write-Host ('Erreur lors de la connexion ou de l''execution : ' + `$_.Exception.Message) -ForegroundColor Red ;
}
"@

                    Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", $escapedCommand
                } else {
                    Write-Host "Numero hors plage, merci de reessayer." -ForegroundColor Red
                }
            } else {
                Write-Host "Entree invalide. Veuillez saisir un numero valide, 'A' pour actualiser ou 'Q' pour quitter" -ForegroundColor Red
            }
        }
    }
} 
