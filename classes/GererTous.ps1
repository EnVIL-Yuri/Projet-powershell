# Importation du fichier CSV
clear
$CSVPath = "\\172.16.122.254\eleves\piartho\SIO2\Powershell\script powershell\classes\2020.csv"

#V�rification du fichier CSV
if (![System.IO.File]::Exists($CSVpath))  
{
  Write-Host "Erreur : le fichier de donnees n'existe pas : ", $CSVpath
  exit   # sortir du script !
}
else
{ Write-Host "le fichier de donnees existe : ", $CSVpath
}

# Charger le fichier CSV en m�moire dans une liste d'objets

$lstCSV = Import-Csv -path $CSVPath -delimiter ";"

# Scan de la liste �l�ment par �l�ment
Write-Host

Foreach(    $ligne IN $lstCSV ) 
{
    $affiche = $ligne.classe
    Write-Host $affiche
    .\GererClasse $affiche
}
