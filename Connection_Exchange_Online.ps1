# =============================
# Script de connexion Exchange Online
# Auteur : @Techparty06 ✨
# =============================

$defaultUsername = 'agent365@mondomaine.fr' # À personnaliser

Clear-Host

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Out-File -Append -FilePath ".\exchange_script.log"
}

# ─────────────────────────────────────────────────────────────
# Configuration : UPN administrateur délégué
# ─────────────────────────────────────────────────────────────
$UsernameInput = Read-Host -Prompt "Entrez l'UPN de l'administrateur délégué ou appuyez sur Entrée pour utiliser le compte par défaut ($defaultUsername)"
$Username = if ([string]::IsNullOrWhiteSpace($UsernameInput)) { $defaultUsername } else { $UsernameInput }

# ─────────────────────────────────────────────────────────────
# Connexion au Partner Center
# ─────────────────────────────────────────────────────────────
Write-Host "`n🔐 Connexion au Centre Partenaire Microsoft 365..." -ForegroundColor Cyan
Write-Host ""
try {
    Connect-PartnerCenter
	Write-Host ""
    Write-Host "✅ Connecté au Centre Partenaire !" -ForegroundColor Green
	Write-Host ""
    Write-Log "Connexion réussie au Partner Center avec $Username"
} catch {
    Write-Host "❌ Échec de la connexion : $($_.Exception.Message)" -ForegroundColor Red
    Write-Log "Erreur de connexion au Partner Center : $($_.Exception.Message)" "ERROR"
    exit
}

function Get-ClientList {
    Write-Host "`n📥 Récupération de la liste des clients..." -ForegroundColor Cyan
	Write-Host ""
    $clients = Get-PartnerCustomer | Sort-Object Name

    $search = Read-Host -Prompt "🔎 Filtrer les clients par nom, domaine ou ID (laisser vide pour tout afficher)"
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
Write-Host "`n🔢 Nombre de clients trouvés : $($clientObj.Count)" -ForegroundColor Magenta

# ─────────────────────────────────────────────────────────────
# Boucle principale
# ─────────────────────────────────────────────────────────────
while ($true) {
    if ($clientObj.Count -eq 0) {
        Write-Host "❌ Aucun client trouvé avec ce filtre. Essayez un autre mot-clé ou tapez 'A'." -ForegroundColor Red
        $clientObj = Get-ClientList
        continue
    }

    Write-Host "`n📋 Liste des tenants disponibles :" -ForegroundColor Yellow
	Write-Host ""
    foreach ($client in $clientObj) {
        $num    = $client.Number.ToString().PadLeft(3)
        $name   = $client.ClientName
        $domain = $client.DomainName
        Write-Host "$num : $name" -ForegroundColor White
        Write-Host "      ↳ $domain" -ForegroundColor DarkGray
    }

    $clientChoice = Read-Host -Prompt "`n👉 Sélectionner le numéro du tenant, taper 'A' pour actualiser, ou 'Q' pour quitter"

    switch ($clientChoice.ToLower()) {
        'q' {
            Write-Host "`n👋 Merci d'avoir utilisé ce script. À bientôt !" -ForegroundColor Cyan
            Write-Log "Script terminé par l'utilisateur"
            exit
        }
        'a' {
            $clientObj = Get-ClientList
			Write-Host "`n🔢 Nombre de clients trouvés : $($clientObj.Count)" -ForegroundColor Magenta
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

                    Write-Host "`n🔄 Connexion à '$ClientName' ($DomainNameId)..." -ForegroundColor Cyan
                    Write-Log "Connexion à $ClientName ($DomainNameId)"

                    $escapedCommand = @"
`$Host.UI.RawUI.WindowTitle = '$windowTitle' ;
try {
    Connect-ExchangeOnline -UserPrincipalName '$Username' -DelegatedOrganization '$DomainNameId' ;
    if (`$?) {
        Write-Host '✅ Connecté à $windowTitle' -ForegroundColor Green ;
        Write-Host ' ' ;
        Write-Host '🔎 Ci-dessous les boîtes aux lettres trouvées :' -ForegroundColor White ;
        Write-Host ' ' ;
        Get-Mailbox -ResultSize Unlimited | Format-Table DisplayName,PrimarySmtpAddress ;
        Write-Host '✨ Tapez votre commande Exchange Online souhaitée' ;
        Write-Host ' ' ;
    } else {
        Write-Host '❌ La connexion a échoué pour $DomainNameId' -ForegroundColor Red ;
    }
} catch {
    Write-Host ('❌ Erreur lors de la connexion ou de l''exécution : ' + `$_.Exception.Message) -ForegroundColor Red ;
}
"@


                    Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", $escapedCommand
                } else {
                    Write-Host "❌ Numéro hors plage, merci de réessayer." -ForegroundColor Red
                }
            } else {
                Write-Host "❗ Entrée invalide. Veuillez saisir un numéro valide, 'A' pour actualiser ou 'Q' pour quitter" -ForegroundColor Red
            }
        }
    }
}
