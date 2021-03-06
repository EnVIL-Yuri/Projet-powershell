clear
$annee = Get-Date -uformat "%Y"

$domaine = "NDLP"
$fqdn = "NDLP.bzh"
$HomeDir = "\\172.19.19.1\eleves" # Dossier de stockage des répertoires personnel
$HomeDrive = "P:" # Lettre du lecteur personnel de l'utilisateur
$LogPath = "C:\sauvegarde_mysql\Creation_Utilisateurs_$date.log"

#Chemin du fchier csv / xls
$chemin = "P:\Mr.Debroise (réseaux)\Si7\maquette powershell\powershell lycee\"+$annee+"\"+$annee+".csv"

if (![System.IO.File]::Exists($chemin)) {
    Write-Host "Le fichier ", $chemin ," n'existe pas."
    exit
} Else {
    Write-Host "Le fichier ", $chemin ," existe."
}

$listCSV = Import-Csv -path $chemin -delimiter ";"

# Scan de la liste élément par élément
Write-Host

Foreach(    $ligne IN $listCSV ) 
{
    $affiche = $ligne.classe
    write-host $affiche
    $chemin1 = "P:\Mr.Debroise (réseaux)\Si7\maquette powershell\powershell lycee\"+$annee+"\"+ $affiche +".csv";
    if (![System.IO.File]::Exists($chemin1)) {
        Write-Host "Le fichier ", $chemin1 ," n'existe pas."
        exit
    } Else {
        Write-Host "Le fichier ", $chemin1 ," existe."
    }

        $listCSV1 = Import-Csv -path $chemin1 -delimiter ";"

        Write-Host "-------------------------------------------------------------";

        Foreach($ligne1 IN $listCSV1 ) 
        {
            $affiche1 = $ligne1.nom
            $affiche2 = $ligne1.prenom
            $displayname = $ligne1.nom+" "+$ligne1.prenom
            $age = $ligne1.ddn
            $mdp = $age
            Write-Host
            Write-Host "-------------------------------------------------------------";
            write-Host
            write-Host $displayname
            write-host "------------"
            If(($nom.Length -lt 4) -and ($prenom.Length -gt 2)) {
			     $nomLength = $nom.Length
			     $nom1 = $nom.Substring(0,$nomLength) 
                 $prenom1 = $prenom.Substring(0,3)
                 $ident = $nom1+$prenom1
			     $ident = $ident.ToLower()
			     $exception = "Exception identifiant pour ",$nom, $prenom," identifiant retenu : ",$ident
			     Write-host "Exception identifiant pour ", $nom ,$prenom , " identifiant retenu : " , $ident >> $LogPath
		    } ElseIf(($prenom.Length -lt 3) -and ($nom.Length -gt 4)) {
			     $prenomLength = $prenom.Length
			     $nom1 = $nom.Substring(0,4)
                 $prenom1 = $prenom.Substring(0,$prenomLength)
                 $ident = $nom1+$prenom1
			     $ident = $ident.ToLower()
			     $exception = "Exception identifiant pour ",$nom, $prenom," identifiant retenu : ",$ident
			     Write-host "Exception identifiant pour ", $nom ,$prenom , " identifiant retenu : " , $ident >> $LogPath
		    } Else {
			     $nom1 = $nom.Substring(0,4)
                 $prenom1 = $prenom.Substring(0,3)
                 $ident = $nom1+$prenom1
			     $ident = $ident.ToLower()
		    }
            $mdp = $mdp.Replace("/","")
            $mdp = $mdp.Remove(4,2)
            
            #$ident contient l'identifiant et $mdp contient le mdp
            write-host $ident
            write-host $mdp
            #New-ADUser -Name $affiche1 -Surname $affiche1 -GivenName $affiche12 -DisplayName $displayname -SamAccountName $ident -AccountPassword (convertto-securestring $mdp -asplaintext -force) -PasswordNeverExpires $true -PasswordNotRequired $true -CannotChangePassword $false -UserPrincipalName ($ident+'@'+$fqdn) -HomeDrive $HomeDrive -HomeDirectory ($HomeDir+'\'+$ident) -Path $NewClasseDN -enabled $true 
    
        }
}