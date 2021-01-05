param($classe)


# Importation du fichier CSV
$CSVPath = "E:\KilianScript\Powershell\"+$classe+".csv"

        Foreach($item IN $lstCSV1 ) 
        {
            $nomcomplet = $item.nom +" "+$item.prenom       
            $nom = $item.nom
            $prenom = $item.prenom
            $age = $item.age
            $pass= $age
            write-host ">>>>>>>>>>>>>>>>>>>>>>>>>>>"
            Write-Host $nomcomplet
            Write-Host $pass
            .\GererUtilisateur $nom $prenom $pass
        }

