# Chemin du fichier CSV
param (
	[string]$xls,
	[string]$csv
)

# Obtention du chemin courant o� est �x�cut� le script
$fullPathIncFileName = $MyInvocation.MyCommand.Definition
$currentScriptName = $MyInvocation.MyCommand.Name
$currentExecutingPath = $fullPathIncFileName.Replace($currentScriptName, "")

# R�cup�ration de la date/heure pour le nom du fichier de log
$date = Get-Date -uformat "%d-%m-%Y_%Hh%Mm"

###################################################################################
# A modifier ! ####################################################################
###################################################################################
$domaine = "NDLP"
$fqdn = "NDLP.bzh"
$HomeDir = "\\172.19.19.1\eleves" # Dossier de stockage des r�pertoires personnel
$HomeDrive = "P:" # Lettre du lecteur personnel de l'utilisateur
$LogPath = "C:\sauvegarde_mysql\Creation_Utilisateurs_$date.log"
###################################################################################
###################################################################################

# Formatage du fichier de log
Write "# V�rification du $date #" > $LogPath
Write "######################################" >> $LogPath
Write " " >> $LogPath

###################################################################################
# Fonctions #######################################################################
###################################################################################

Function NewUser
{
# Debug-Pat
Write-Host  "Debug : pass = " $pass

	# Sinon cr�ation du nouvel utilisateur
	New-ADUser -Name $displayname -Surname $nom -GivenName $prenom -DisplayName $displayname -SamAccountName $login -AccountPassword (convertto-securestring $pass -asplaintext -force) -PasswordNeverExpires $true -PasswordNotRequired $true -CannotChangePassword $false -UserPrincipalName ($login+'@'+$fqdn) -HomeDrive $HomeDrive -HomeDirectory ($HomeDir+'\'+$login) -Path $NewClasseDN -enabled $true 
	$UserDNsearch = Get-ADUser -Filter {SamAccountName -eq $login}
	# Ajout au groupe
	$NewGroupDNsearch = Get-ADGroup -Filter {Name -eq $NewClasse}
	$NewGroup = $NewGroupDNsearch.distinguishedName # CN=NewGroup,CN=Users,DC=domain,DC=local
	$GroupDest = [ADSI]"LDAP://$NewGroup"
	$GroupDestDN = "LDAP://$UserDNsearch"
	$GroupDest.Add($GroupDestDN) # Ajout de l'utilisateur au nouveau groupe
	# Creation du home directory
	$UserHomeDir = New-Item ($HomeDir+'\'+$login) -type directory
Write-Host  "Debug : UserHomeDir" $UserHomeDir
	# Mise en place des ACL du home directory
	$account = $domaine+'\'+$login
	$droits = [System.Security.AccessControl.FileSystemRights]::FullControl
	$inheritance = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
	$allowdeny = [System.Security.AccessControl.AccessControlType]::Allow
	$propagation = [System.Security.AccessControl.PropagationFlags]::None
	$dirACE = New-Object System.Security.AccessControl.FileSystemAccessRule($account,$droits,$inheritance,$propagation,$allowdeny)
Write-Host  "Debug : dirACE = " $dirACE
	$dirACL = Get-Acl $UserHomeDir
Write-Host  "Debug : dirACL = " $dirACL
	$dirACL.AddAccessRule($dirACE)
	Set-Acl $UserHomeDir $dirACL
	Write-Host -ForegroundColor Green "$displayname ($login) � �t� cr�� dans $NewClasseDN.`r"
	Write-Host -ForegroundColor Green "$displayname ($login) � �t� ajout� au groupe $NewGroup`r"
	Write-Host -ForegroundColor Green "Droits sur $UserHomeDir assign�s.`r"
	Write-Host ""
	Write "$displayname ($login) � �t� cr�� dans $NewClasseDN.`r"    >> $LogPath
	Write "$displayname ($login) � �t� ajout� au groupe $NewGroup`r" >> $LogPath
	Write "Droits sur $UserHomeDir assign�s.`r" >> $LogPath	
	Write " " >> $LogPath
	[console]::ResetColor()
}


Function CreateUser
{
	[console]::ForegroundColor = "yellow"
	$CSVCorrect = Read-Host "Le fichier CSV vous semble-t-il correct ? (Y/N)"
	If(($CSVCorrect -eq "N") -or ($CSVCorrect -eq "n")) {
		Write-Host ""
		Write-Host -Foregroundcolor Red "Veuillez apporter les modifications au fichier et recommencer l'op�ration."
		Write-Host -Foregroundcolor Red "Lors de la prochaine �x�cution sp�cifier le parametre -csv NomDuCSVmodifie.csv � la place du param�tre -xls."
		Write-Host " "
		Write-Host -Foregroundcolor Red "Exemple: ./Create-ADUser.ps1 -csv NomDuCSVmodifier.csv"
		Break
	} Else {
	
	#############################################################
	# Cr�ation des utilisateurs #################################
	#############################################################

	# Chargement du module Active Directory
	Import-Module ActiveDirectory
	Clear
	[console]::ResetColor()
	
	# affichage du fichier CSV final
	Import-CSV -path $CSVPath -delimiter ";" | Format-Table 
	
	[console]::ForegroundColor = "yellow"
	Write-Host -Foregroundcolor Red "/!\ Derni�re v�rification avant cr�ation !"
	Write-Host ""
	Read-Host "Pressez une touche pour valider ces informations...ou CTRL+C pour annuler "
	[console]::ResetColor()
	
	$users = Import-CSV -path $CSVPath -delimiter ";"
	
	Foreach($user IN $users) {
	# R�cup�ration des informations du CSV
    $pass = $user.Naissance
    $nom = $user.Nom
    $prenom = $user.Prenom
    $displayname = $user.Nom +" "+ $user.Prenom
    $login = $user.login
	$NewClasse = $user.Classe		
	$NewClasseDN = (Get-ADOrganizationalUnit -Filter {Name -eq $NewClasse}).DistinguishedName # OU=OUdestination,DC=domain,DC=local
	Write-Host -ForegroundColor Yellow "$displayname ($login) en cours...`r"
	Write "$displayname ($login) en cours..." >> $LogPath
	# D�but des v�rifications
	# Si l'utilisateur existe...
	If(Get-ADUser -Filter {SamAccountName -eq $login}) {
		$UserDNsearch = Get-ADUser $login
		[console]::ForegroundColor = "cyan"
		$ChangeUser = Read-Host "$($UserDNsearch.Name) ($($UserDNsearch.SamAccountName)) poss�de d�ja ce login, souhaitez-vous changer l'identifiant de $displayname ? (Y/N)"
		Write "$($UserDNsearch.Name) ($($UserDNsearch.SamAccountName)) poss�de d�ja ce login, souhaitez-vous changer l'identifiant de $displayname ? (Y/N)" >> $LogPath
		If(($ChangeUser -eq "Y") -or ($ChangeUser -eq "y")) {
			Write "Reponse = YES`r" >> $LogPath
			[console]::ForegroundColor = "yellow"
			$SetSAM = Read-Host "Nouvel identifiant ?"
			[console]::ResetColor()
			$login = $SetSAM
			# fonction NewUser
			NewUser $login
		} Else {
		Write "Reponse = NO`r" >> $LogPath
		$UserDN = $UserDNsearch.distinguishedName # CN=DUPONT Eric,OU=OUactuelle,DC=domain,DC=local
		$UserName = $UserDNsearch.Name
		$AncClasse = $UserDN.Replace("CN=$UserName,","") # Supression du CN => OU=OUactuelle,DC=domain,DC=local
		# Si l'utilisateur doit changer de classe
		If(!($NewClasseDN -eq $AncClasse)) {
			$con = [ADSI]"LDAP://$AncClasse"
			$ANCGroup = $con.PSBase.Children.Find("CN=$UserName") # Cherche le groupe actuel de l'utilisateur
			$AncGroup = $ANCGroup.memberof # CN=UserGroup,CN=Users,DC=domain,DC=local
			$NewGroup = (Get-ADGroup -Filter {Name -eq $NewClasse}).DistinguishedName # CN=NewGroup,CN=Users,DC=domain,DC=local
			$GroupDest = [ADSI]"LDAP://$NewGroup"
			$GroupDestDN = "LDAP://$UserDNsearch"
			$GroupAdd = $GroupDest.Add($GroupDestDN) # Ajout de l'utilisateur au nouveau groupe
			$GroupAnc = [ADSI]"LDAP://$AncGroup"
			$GroupRemove = $GroupAnc.Remove($GroupDestDN) # Suppression de l'ancien groupe
			# OU 
			$OUDest = [ADSI]"LDAP://$NewClasseDN" # Connexion a l'OU de destination
			$OUDest.MoveHere("LDAP://CN=$UserName,$AncClasse","CN=$UserName") # Deplace l'utilisateur depuis son ancienne OU
			Write-Host -ForegroundColor Green "$UserName � �t� d�pla�� de $AncClasse vers $NewClasseDN`r"
			Write-Host -ForegroundColor Green "$UserName � �t� d�pla�� du groupe $AncGroup vers $NewGroup`r"
			Write-Host ""
			Write "$UserName � �t� d�pla�� du groupe $AncClasse vers $NewClasseDN`r" >> $LogPath
			Write "$UserName � �t� d�pla�� du groupe $AncGroup vers $NewGroup`r" >> $LogPath
			Write " " >> $LogPath
			[console]::ResetColor()
		} Else {
			# Utilisateur redoublant
			Write-Host -ForegroundColor Green "$UserName existe d�ja et redouble.`r"
			Write-Host ""
			Write "$UserName existe d�ja et redouble.`r" >> $LogPath
			Write " " >> $LogPath
			[console]::ResetColor()
		}
		}	
	} Else {
		# Si l'utilisateur n'existe pas => cr�ation
		NewUser
	}
}
}
}

###################################################################################
# D�but du script #################################################################
###################################################################################

Clear

# V�rification si parametre CSV est sp�cifi�
If($csv -eq "") {

	# Fichier CSV de sortie apr�s conversion
	$CSVPath = $currentExecutingPath + "convert.csv"

	# Convertion du fichier .xls en .csv
	$xlCSV=6
	$Excelfilename = "$xls"
	$CSVfilename = "$CSVPath"
	$Excel = New-Object -comobject Excel.Application
	$Excel.Visible = $False
	$Excel.displayalerts=$False
	$Workbook = $Excel.Workbooks.Open($ExcelFileName)
Write-Host "Debug : CSVfilename " $CSVfilename
Write-Host "Debug : xlCSV       " $xlCSV
	$Workbook.SaveAs($CSVfilename,$xlCSV)
	$Excel.Quit()
	If(ps excel){kill -name excel}

	# Fonction supression accent
	Function Remove-Diacritics([string]$String)
	{
		$objD = $String.Normalize([Text.NormalizationForm]::FormD)
		$sb = New-Object Text.StringBuilder
	 
		for ($i = 0; $i -lt $objD.Length; $i++) {
			$c = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($objD[$i])

			if($c -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
				# Test ajout� par Pat
#				if ( $objD[$i] -eq '�' ) {
				  [void]$sb.Append($objD[$i])
#				}
			}
		  }
	 
		return("$sb".Normalize([Text.NormalizationForm]::FormD))
	}

	# Convertion du CSV Unicode en UTF8
	$CSVFile = Get-Content $CSVPath
	$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($False)
	[System.IO.File]::WriteAllLines($CSVPath, $CSVFile, $Utf8NoBomEncoding)

	Write-Host -Foregroundcolor Yellow "Conversion du fichier CSV (Unicode => UTF8) en cours..."
	Write-Host -Foregroundcolor Yellow "Traitement du fichier CSV en cours..."

	# Importation du fichier CSV Aplon
	$ImportCSV = Import-CSV -path "$CSVPath" -delimiter ";"

	# Supression du fichier de conversion xls => csv
	Remove-Item -Path $CSVPath
	IF(Test-Path $CSVPath){
		Write-Host -ForegroundColor Red "Erreur lors de la supression du fichier $CSVPath"
	}

	# Formatage fichier CSV Aplon
	Foreach($item IN $ImportCSV) {
		# Nom
		$nom = $item.Nom
		$nom = $nom.Replace(" ","")
		$nom = Remove-Diacritics([string]$nom)
		# Pr�nom
		$prenom = $item.Prenom
		$prenom = $prenom.Replace(" ","")
		$prenom = $prenom.Replace("�","e")
		$prenom = Remove-Diacritics([string]$prenom)
		# Formatage du mot de passe - ex: 01/02/1990 => 010290
		$password = $item.Naissance
		$password = $password.Replace("/","")
		$password = $password.Remove(4,2)
		# Classe
		$classe = $item.Classe
		$classe = $classe.Replace(" ","")
		$item | Add-Member -memberType noteProperty -name Nom       -value "$nom"      -force
		$item | Add-Member -memberType noteProperty -name Prenom    -value "$prenom"   -force
		$item | Add-Member -memberType noteProperty -name Naissance -value "$password" -force
		$item | Add-Member -memberType noteProperty -name Classe    -value "$classe"   -force
		# login
		$prenom = $prenom.Replace("-","")
		$nom = $nom.Replace("-","")
		$login = $item.login
		If(($nom.Length -lt 4) -and ($prenom.Length -gt 2)) {
			$nomLength = $nom.Length
			$login = $nom.SubString(0,$nomLength) + $prenom.SubString(0,3)
			$login = $login.ToLower()
			$exception = "Exception identifiant pour $nom $prenom, identifiant retenu : $login"
			Write "Exception identifiant pour $nom $prenom, identifiant retenu : $login" >> $LogPath
		} 
		ElseIf(($prenom.Length -lt 3) -and ($nom.Length -gt 4)) {
			$prenomLength = $prenom.Length
			$login = $nom.SubString(0,4) + $prenom.SubString(0,$prenomLength)
			$login = $login.ToLower()
			$exception = "Exception identifiant pour $nom $prenom, identifiant retenu : $login"
			Write "Exception identifiant pour $nom $prenom, identifiant retenu : $login" >> $LogPath
		}
		Else {
			$login = $nom.SubString(0,4) + $prenom.SubString(0,3)
			$login = $login.ToLower()
		}
		$item | Add-Member -memberType noteProperty -name login -value "$login" -force
	}

	$ImportCSV | Export-CSV $CSVPath -NoTypeInformation -Encoding "UTF8" -delimiter ";"

	# affichage du fichier CSV final
	Import-CSV -path $CSVPath -delimiter ";" | Format-Table

	Write-Host ""
	Write-Host -ForegroundColor Cyan $exception
	Write-Host ""
	# Appel de la fonction de cr�ation des utilisateurs
	CreateUser $CSVPath
} Else {
	$CSVPath = Read-Host "Quel est le chemin du fichier CSV ?"
	# Appel de la fonction de cr�ation des utilisateurs
	CreateUser $CSVPath
}