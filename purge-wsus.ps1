#* Alex Cachon 14/12/2021
#*Suppression auto des postes dans WSUS

param(
    [Switch] $admin
)
function Admin-Mode 
{
    if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) 
    { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
}

function Purge_Wsus{
# Import des modules et assembly nécessaire pour l'execution du script
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")

# Déclaration des variables
$Counter = 0
# Déclaraton des serveurs wsus
$WsusINTArmor = "ASP-INF-WSUS01.INT01.ARMORSANTE.ORG"
$WsusPaimpol = "PLP-INF-WSUS05.INT01.ARMORSANTE.ORG"
$WsusTreguier = "TGP-INF-WSUS04.INT01.ARMORSANTE.ORG"
$WsusGuingamp = "GPP-INF-WSUS01.INT01.ARMORSANTE.ORG"
$WsusLannion = "LNP-INF-WSUS01.INT01.ARMORSANTE.ORG"
$WsusLannion2 = "LNP-INF-WSUS02.INT01.ARMORSANTE.ORG"
$WsusMazier = "MZP-INF-WSUS01.INT01.ARMORSANTE.ORG"
# Creation du fichier de log
$Date = Get-Date -Format "ddMMyyyy-HHmmss"
$folderoutput = "C:\Scripts\Purge_computer\logs\"
$file_output = "C:\Scripts\Purge_computer\logs\del_computer-"+$Date+".csv"
# Définition des noms de domain
$INT = "int01.armorsante.org"
$SB = "ch-stbrieuc.fr"
$GP = "ch-guingamp.info"
$TG = "ch-treguier.int"
$PL = "hexagone.int"
$LN = "ch-lannion.fr"

# Exclusion Poste
[array]$exclusions = @("asp-ldap-dc02","srvgescle","mpl2vmsil-sos","siasrv01","mpl2vmsil","asp-inf-dl01","asp-web-wap02","siasrv02","srvttdc01","srvretino","asp-ldap-dc01","srvca01","logifsiweb","asp-web-wap01","srvmplevo2")
Add-Content -Path $file_output "Machine/Domain"
# Connection aux serveurs Wsus

                $WsusINTArmor = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($WsusINTArmor,$true,8531)
                $WsusPaimpol = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($WsusPaimpol,$true,8531)
                $WsusTreguier = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($WsusTreguier,$true,8531)
                $WsusLannion = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($WsusLannion,$true,8531)
                $WsusLannion2 = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($WsusLannion2,$true,8531)
                $WsusMazier = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($WsusMazier,$true,8531)
                [array]$AllWsus = @($WsusINTArmor,$WsusPaimpol,$WsusTreguier,$WsusGuingamp,$WsusLannion,$WsusLannion2,$WsusMazier)        

# Connection au serveur Wsus Esclave (si besoin)

# Obtention du path utilisateur temporaire
$TempPath = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath "CheckWsus"
# Vérification du fichier temporaire existant
if(Test-Path -path $TempPath)
{
Remove-Item -Path $TempPath -Recurse -Force
}
 New-Item -Path $TempPath -ItemType Directory -Force # Création d'un dossier temporaire (dans le C:\Users\%user%\AppData\Local\Temp\)
 Clear-Host


# Génération des fichiers temporaires
$TempFile = [System.Guid]::NewGuid().Guid + ".txt" 
$TempFileWsus = [System.Guid]::NewGuid().Guid + ".txt" # Génère un nom aléatoirement avec guid
$FullTempFilePath = Join-Path -Path $TempPath -ChildPath $TempFile # Place ce fichier dans le dossier précédement créer
$FullTempFilePathWsus = Join-Path -Path $TempPath -ChildPath $TempFileWsus

# Récupération des PC dans WSUS
$WsusINTArmor | Get-WsusComputer -IncludeDownstreamComputerTargets | Select-Object FullDomainName |Out-File $FullTempFilePath
# Modification du fichier Wsus -> garder seulement les noms de postes
$WSDM = Get-Content $FullTempFilePath
$WSDM[3..($WSDM.length-1)]|Out-File $FullTempFilePath -Force
$WSDM | Where-Object {$_.trim() -ne "" } | Out-File $FullTempFilePath -Force
$WSDM = Get-Content $FullTempFilePath

foreach ($PC in $WSDM)
{
    #Filtre les noms de domain
    if($PC -match $INT -or $PC -match $SB -or $PC -match $GP -or $PC -match $TG -or $PC -match $PL -or $PC -match $LN -or $PC -match ".armorsante.org"-or $PC -match ".chsb.fr" -or $PC -match ".ght22.bzh" -or $PC -match "ch-guingamp.bloc")
    {
    $PCedit = $PC.substring(0,$PC.IndexOf('.'))
    $PCedit -replace ' ',''| Out-File $FullTempFilePathWsus -Append
    }
    else 
    {
    $PCedit = $PC -replace ' ',''
    $PCedit | Out-File $FullTempFilePathWsus -Append
    }
  }

$NFWSUS = Get-Content $FullTempFilePathWsus
$NFWSUS | Select-Object -index (3..$NFWSUS.Length) | Out-File $FullTempFilePathWsus
Import-Csv $FullTempFilePathWsus | Sort-Object Machine –Unique # Supprime les doublons

# Suppression des postes dans le fichier si ceux-ci sont dans la liste d'exclusion
$NFWSUS = Get-Content $FullTempFilePathWsus
foreach ($PC in $NFWSUS)
{
foreach ($exclusion in $exclusions)
{
if ($PC -match $exclusion) {
    Write-Host $PC " est dans les exclusions, celui-ci ne sera pas traite"
    (Get-Content $FullTempFilePathWsus) -notmatch $PC | Out-File $FullTempFilePathWsus
}
}
}

# Obtention du path utilisateur RAW dans le répertoire temporaire wsuscheck
$TempPathraw = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath "CheckWsus\RAW"
# Vérification du fichier temporaire existant
if(Test-Path -path $TempPathraw)
{
Remove-Item -Path $TempPathraw -Recurse -Force
}
 New-Item -Path $TempPathraw -ItemType Directory -Force # Création d'un dossier temporaire (dans le C:\Users\%user%\AppData\Local\Temp\)
Move-Item -Path $FullTempFilePath -Destination $TempPathraw # Déplace le fichier originel pour pouvoir concaténé sans obtenir les ressources de celui-ci
$NewTempFilePath = $TempPathraw+"\"+$TempFile # Déclaration du nouvelle emplacement

start-sleep 3 
$ScrapWsusPCs  = get-content $FullTempFilePathWsus

# Bouclage sur chaque occurence pour check les ordinateurs dans l'AD)
$Counter = 0
Clear-Host
foreach ($ScrapWsusPC in $ScrapWsusPCs) {
$P = ((($Counter++) / $ScrapWsusPCs.Count) * 100)
Write-Progress -Activity "Check des PC et Serveurs dans l'AD" -Status 'Progression :' -CurrentOperation $P"%" -PercentComplete $P
#Recherche le Poste sur tout les domains
try {
    $notview = Get-ADComputer -identity $ScrapWsusPC -Server $INT
}
catch {
    try {
        $notview = Get-ADComputer -identity $ScrapWsusPC -Server $PL
    }
    catch {
        try {
            $notview = Get-ADComputer -identity $ScrapWsusPC -Server $SB
        }
        catch {
            try {
                $notview =  Get-ADComputer -identity $ScrapWsusPC -Server $GP
            }
            catch {
                try {
                    $notview =  Get-ADComputer -identity $ScrapWsusPC -Server $LN
                }
                catch {
                   try {
                    $notview = Get-ADComputer -identity $ScrapWsusPC -Server $TG
                   }
                   catch {
                       #recherche et récupération dans le fichier original le FullDomainName
                       foreach ($line in $NewTempFilePath)
                       {
                        $EditScrapWsusPC = Select-String -path $NewTempFilePath -pattern $ScrapWsusPC #recherche du poste dans le fichier
                        $EditScrapWsusPC = $EditScrapWsusPC.ToString() #transformation en format texte pour faire des actions dessus
                        $EditScrapWsusPC = $EditScrapWsusPC.Split(':')[$($EditScrapWsusPC.Split(':').count-1)] # Enlève les informations inutiles pour garder que le fulldomainname
                        $DCGet = $EditScrapWsusPC -replace $ScrapWsusPC,''
                       }
                    write-host $ScrapWsusPC "not found and delete"
                    #Suppression de la machine dans tout les wsus
                    foreach ($wsus in $AllWsus)
                    {
                        try {
                                               $client = $wsus.SearchComputerTargets($EditScrapWsusPC)
                                               #$client.delete() 
                        }
                        catch {}
                    }
                    # Génération dus fichier temporaires des PC/serveurs non trouvé
                    $TempFile = [System.Guid]::NewGuid().Guid + ".txt"
                    $TempFilepc = [System.Guid]::NewGuid().Guid + ".txt"
                    $FullTempFilePath = Join-Path -Path $TempPath -ChildPath $TempFile
                    $TempFilepc = Join-Path -Path $TempPath -ChildPath $TempFilepc
                    #export le nom du pc vers fichier temporaire
                    $ScrapWsusPC+"/"+$DCGet | Out-File $FullTempFilePath -Force
                }
            } 
          }
        }
    }
  }
}
        # Suppression du fichier des postes 
        Remove-Item  $FullTempFilePathWsus
        # Concaténation des fichiers temporaire des postes supprimer vers le fichier définitif
        $TempFiles = Get-ChildItem -Path $TempPath
        $Counter = 0
        foreach ($TempFile in $TempFiles) {
            $C = ((($Counter++) / $TempFiles.Count) * 100)
            Write-Progress -Activity 'Concaténation des fichiers' -Status 'Progression :' -CurrentOperation $C"%" -PercentComplete $C
            Get-Content -Path $TempFile.FullName | Out-File $file_output -Append #Concaténation des fichiers
        }
        # Suppression des données temporaires
        Remove-Item -Path $TempPath -Force -Recurse 
        #>
function xlsx{
# Défini la location et les délimiteurs
$csv = $file_output #Fichier source
$xlsx = $folderoutput+"del_computer-"+$Date+".xlsx" #fichier destination
$delimiter = "/" #délimiteur du fichier csv

# Créer un nouveau classeur Excel avec une feuille vide
$excel = New-Object -ComObject excel.application 
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)

# Générez la commande QueryTables.Add et reformatez les données
$TxtConnector = ("TEXT;" + $csv)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

# Exécuter et supprimer la requête d'importation
$query.Refresh()
$query.Delete()

# Enregistrez et fermez le classeur en tant que XLSX et supprimez csv.
$excel.DisplayAlerts = $false;
$Workbook.SaveAs($xlsx,51)
$excel.Quit()
}

# Conversion du fichier csv en xlsx
xlsx
# Suppression du csv
Remove-Item $file_output

# Envoie du mail        
$body1= "voici les postes supprimes ce jour $($Date)"
$body = @($body1
) |Out-String
#! modifier l'adresse mail après test par "supervision.dsi@armorsante.bzh"
Send-MailMessage -SmtpServer "smtp.armorsante.bzh" -From "delete.computer@armorsante.bzh" -To "alex.cachon.ext@armorsante.bzh" -Subject "suppression WSUS" -Body $body -Attachments $xlsx
}
#lancement en mode admin (arguement -admin)
if ($admin -eq $true){
#Execution de la fontion Admin-mode
Admin-Mode
}
#Execution de la fontion purge wsus
Purge_Wsus