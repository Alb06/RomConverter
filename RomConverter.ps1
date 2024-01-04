Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#gestion des logs
function Get-LogFile {
    # Obtenir le chemin du script en cours d'exécution
    $scriptPath = $PSCommandPath
    # Construire le chemin du fichier de log en changeant l'extension
    $logFile = [IO.Path]::ChangeExtension($scriptPath, ".log")

    # Vérifier si le fichier de log existe déjà
    if (-not (Test-Path $logFile)) {
        # Créer un nouveau fichier de log vide
        New-Item -Path $logFile -ItemType File
    }
    return $logFile
}
function Write-Log {
    param (
        [string]$logFile,
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if (-not (Test-Path $logFile)) {
        # Créer et récupérer un nouveau fichier de log vide
        $logFile = Get-LogFile
    }

    if ($Message -is [System.Collections.IEnumerable]) {
        foreach ($line in $Message) {
            $logMessage = "$timestamp - $line"
            Add-Content -Path $logFile -Value $logMessage -Encoding UTF8
        }
    }
    else {
        $logMessage = "$timestamp - $Message"
        Add-Content -Path $logFile -Value $logMessage -Encoding UTF8
    }
}

# Assurez-vous que 7-Zip est installé
function Get-7ZipPath {
    param (
        [string]$logFile
    )

    Write-Log $logFile "Recherche du chemin d'installation de 7-Zip."
    try {
        $key = Get-ItemProperty -Path "HKLM:\SOFTWARE\7-Zip" -ErrorAction Stop
        if ($key) {
            Write-Log $logFile "7-Zip a bien été trouvé sur la machine."
            return Join-Path $key.Path "7z.exe"
        }
    } catch {
        Write-Log $logFile "7-Zip n'est pas installé sur la machine."
        $response = [System.Windows.Forms.MessageBox]::Show("7-Zip n'est pas installé. Voulez-vous l'installer maintenant ?", "Installer 7-Zip", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log $logFile "L'utilisateur a décider d'installer 7-Zip sur sa machine."
            Start-Process "https://www.7-zip.org/download.html"
        }

        Write-Log $logFile "L'utilisateur a décider de ne pas installer 7-Zip sur sa machine."
        Write-Warning "7-Zip n'est pas installé ou le chemin n'a pas pu être trouvé."
        exit
    }
    return $null
}

#Vérification de la présence de chdman.exe
function Verify-Chdman {
    param (
        [string]$logFile
    )

    $chdmanPath = Join-Path $PSScriptRoot "chdman.exe"
    if (-not (Test-Path $chdmanPath)) {
        Write-Log $logFile "Chdman.exe n'est pas dans le même dossier que ce script"
        $response = [System.Windows.Forms.MessageBox]::Show("Le fichier 'chdman.exe' est nécessaire pour la conversion des fichiers et doit se trouver dans le même dossier que ce script. Voulez-vous télécharger 'chdman.exe' maintenant ?", "Télécharger chdman.exe", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log $logFile "Lancement du téléchargement de Chdman.exe par l'utilisateur"
            Start-Process "https://wiki.recalbox.com/tutorials/utilities/rom-conversion/chdman/chdman.zip"
            exit
        } else {
            Write-Log $logFile "L'utilisateur ne souhaite pas télécharger Chdman.exe"
            Write-Warning "chdman.exe est nécessaire pour continuer. Veuillez placer chdman.exe dans le même dossier que ce script."
            exit
        }
    }
}

#Selection des dossiers de travail
function Select-Folder([string]$logFile, $description, $rootFolder, $checkType) {
    do {
        Write-Log $logFile "Affichage de la boîte de dialogue pour sélectionner le dossier : $description"
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = $description
        if ($rootFolder) { $folderBrowser.SelectedPath = $rootFolder }

        $result = $folderBrowser.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            Write-Log $logFile "Sélection de dossier annulée par l'utilisateur."
            exit
        }

        $selectedPath = $folderBrowser.SelectedPath
        $isValid = $false

        # Vérifier si le chemin contient des caractères spéciaux
        $matches = [regex]::Matches($selectedPath, "[^a-zA-Z0-9\\s\-_: ]")
        if ($matches.Count -gt 0) {
            $invalidChars = $matches | ForEach-Object { $_.Value } | Sort-Object -Unique
            Write-Log $logFile "Le chemin sélectionné contient des caractères spéciaux : $invalidChars"
            $isValid = $false
            $errorMessage = "Le chemin du dossier ne doit pas contenir de caractères spéciaux. `r`nChemin proposé : $selectedPath `r`nCaractères problématiques : $($invalidChars -join ', ')"
        } 
        else {
            switch ($checkType) {
                "7zSource" {
                    $isValid = (Get-ChildItem -Path $selectedPath -Filter *.7z -File | Measure-Object).Count -gt 0
                    $errorMessage = "Aucun fichier .7z trouvé dans le dossier sélectionné."
                }
                "Empty" {
                    $filesInFolder = Get-ChildItem -Path $selectedPath -File
                    $isValid = ($filesInFolder.Count -eq 0)
                    $errorMessage = "Le dossier sélectionné doit être vide."
                }
                "CHD" {
                    $isValid = (Get-ChildItem -Path $selectedPath -Filter *.chd -File | Measure-Object).Count -gt 0 -or (Get-ChildItem -Path $selectedPath | Measure-Object).Count -eq 0
                    $errorMessage = "Le dossier sélectionné doit être vide ou contenir uniquement des fichiers .chd."
                }
                "7zOutput" {
                    $isValid = (Get-ChildItem -Path $selectedPath -Filter *.7z -File | Measure-Object).Count -gt 0 -or (Get-ChildItem -Path $selectedPath | Measure-Object).Count -eq 0
                    $errorMessage = "Le dossier sélectionné doit être vide ou contenir uniquement des fichiers .7z."
                }
            }
        }

        if (-not $isValid) {
            Write-Log $logFile "Selection du dossier par l'utilisateur invalide"
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Sélection invalide", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } while (-not $isValid)

    return $selectedPath
}

#chargement des parametres
function Load-Settings {
    param (
        [string]$logFile
    )

    Write-Log $logFile "Chargement des paramètres depuis le fichier de configuration."
    $settingsFile = Join-Path $PSScriptRoot "settings.txt"
    if (Test-Path $settingsFile) {
        Get-Content $settingsFile
    } else {
        "", "", "", ""
    }
}

#sauvgarde des parametres
function Save-Settings([string]$logFile, $sourceFolder, $decompressionFolder, $chdFolder, $recompressionFolder) {
    Write-Log $logFile "Enregistrement des paramètres dans le fichier de configuration."
    $settingsFile = Join-Path $PSScriptRoot "settings.txt"
    Set-Content -Path $settingsFile -Value @("$sourceFolder", "$decompressionFolder", "$chdFolder", "$recompressionFolder")
}

function Close-And-Dispose-All {
    Write-Host "Le ScriptBlock a terminé son exécution."
    Write-Host "Appuyez sur une touche pour fermer la fenêtre..."
    # Attendre l'appui sur une touche
    [Console]::ReadKey()

    $runspace.Close()
    $runspace.Dispose()
    $powerShell.Dispose()
    $progressUpdates.Dispose()
    $timer.Dispose()
    $pauseEvent.Dispose()

    $form.Close()
}

# Chemin vers les dossiers necessaires
$logFile =              Get-LogFile
$previousSettings =     Load-Settings $logFile
$sourceFolder =         Select-Folder $logFile "Sélectionnez le dossier source des fichiers .7z" $previousSettings[0] "7zSource"
$decompressionFolder =  Select-Folder $logFile "Sélectionnez le dossier de decompression" $previousSettings[1] "Empty"
$chdFolder =            Select-Folder $logFile "Sélectionnez le dossier de stockage .chd" $previousSettings[2] "CHD"
$recompressionFolder =  Select-Folder $logFile "Sélectionnez le dossier de recompression" $previousSettings[3] "7zOutput"
Save-Settings $logFile "$sourceFolder" "$decompressionFolder" "$chdFolder" "$recompressionFolder"
$pathTo7Zip =           Get-7ZipPath $logFile
Verify-Chdman $logFile

# Créer un ManualResetEvent pour gérer la pause
$pauseEvent = New-Object System.Threading.ManualResetEvent($true)

# ScriptBlock pour la décompression
$scriptBlock = {
    param ($sourceFolder, $decompressionFolder, $chdFolder, $pathTo7Zip, $pathToScript, $pauseEvent, [System.Collections.Concurrent.BlockingCollection[int]]$progressUpdates, [System.Collections.Concurrent.BlockingCollection[string]]$outputUpdates)

    # Obtenir tous les fichiers .7z dans le dossier source
    $files = Get-ChildItem -Path $sourceFolder -Filter *.7z
    $totalFiles = $files.Count
    $processedFiles = 0

    foreach ($file in $files) {
        # Attendre que le ManualResetEvent soit signalé (non pausé)
        $pauseEvent.WaitOne()

        # Traitement des fichiers .7z vers décompressés
        $currentFile = $file.FullName
        $outputUpdates.Add("7ZIP start to process on " + $file.Name)
        $output = & $pathTo7Zip x $currentFile -o"$decompressionFolder" 2>&1
        $output | ForEach-Object {
            $outputUpdates.Add($_)
        }

        # Traitement des fichiers décompressés vers .chd
        $unzippedFiles = Get-ChildItem -Path $decompressionFolder -Recurse -Include *.cue, *.gdi, *.iso
        $outputCHD = Join-Path $chdFolder ($file.BaseName + ".chd")
        foreach ($unzippedFile in $unzippedFiles) {
            $currentFile = $unzippedFile.FullName
            $outputUpdates.Add("CHDMAN start to process on " + $unzippedFile.Name)
            $output = & "$pathToScript\chdman.exe" createcd -i "$currentFile" -o "$outputCHD" 2>&1
            $output | ForEach-Object {
                $outputUpdates.Add($_)
            }
        }

        # Nettoyage du dossier de décompression
        Remove-Item $decompressionFolder\* -Force

        # Mettre à jour la progression
        $processedFiles++
        $progress = [math]::Round(($processedFiles / $totalFiles) * 100)
        $progressUpdates.Add($progress)
    }

    # Signaler que l'ajout d'éléments est terminé
    $progressUpdates.CompleteAdding()
    $outputUpdates.CompleteAdding()
}

# Créer un Runspace
$runspace = [runspacefactory]::CreateRunspace()
$runspace.Open()

# Créer une BlockingCollection pour les mises à jour de progression
$progressUpdates = New-Object System.Collections.Concurrent.BlockingCollection[int]
$outputUpdates = New-Object System.Collections.Concurrent.BlockingCollection[string]

# Créer la fenêtre de l'interface utilisateur
$form = New-Object System.Windows.Forms.Form
$form.Text = "Décompression des fichiers"
$form.Size = New-Object System.Drawing.Size(400, 150)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 10)
$progressBar.Size = New-Object System.Drawing.Size(360, 23)
$progressBar.Style = 'Continuous'
$form.Controls.Add($progressBar)

# Ajouter un bouton pour mettre en pause/continuer
$pauseButton = New-Object System.Windows.Forms.Button
$pauseButton.Location = New-Object System.Drawing.Point(10, 40)
$pauseButton.Size = New-Object System.Drawing.Size(360, 23)
$pauseButton.Text = "Pause"
$pauseButton.Add_Click({
    if ($pauseEvent.WaitOne(0)) {
        $pauseEvent.Reset() # Mettre en pause
        $pauseButton.Text = "Continuer"
    } else {
        $pauseEvent.Set() # Reprendre
        $pauseButton.Text = "Pause"
    }
})
$form.Controls.Add($pauseButton)

# Créer et démarrer le PowerShell avec le ScriptBlock
$powerShell = [powershell]::Create().AddScript($scriptBlock).AddArgument($sourceFolder).AddArgument($decompressionFolder).AddArgument($chdFolder).AddArgument($pathTo7Zip).AddArgument($PSScriptRoot).AddArgument($pauseEvent).AddArgument($progressUpdates).AddArgument($outputUpdates)
$powerShell.Runspace = $runspace
$asyncResult = $powerShell.BeginInvoke()

# Ajouter un Timer pour mettre à jour la barre de progression et afficher la sortie
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 100
$timer.Add_Tick({
    while ($progressUpdates.Count -gt 0) {
        $progressBar.Value = $progressUpdates.Take()
        if($progressUpdates.IsAddingCompleted) {
            break
        }
    }

    while ($outputUpdates.Count -gt 0) {
        $output = $outputUpdates.Take()
        # Gestion Console et Logs
        if($output -match "start to process on") {
            Write-Host $output -ForegroundColor Blue
            Write-Log $logFile $output
        }
        if($output -match "(?i)error") {
            Write-Host "  $output" -ForegroundColor Red
            Write-Log $logFile "  $output"
        }
        else {
            Write-Host "  $output"
            Write-Log $logFile "  $output"
        }
        if($outputUpdates.IsAddingCompleted) {
            break
        }
    }
    if($progressUpdates.IsAddingCompleted -and $outputUpdates.IsAddingCompleted) {
        Close-And-Dispose-All
    }
})
$timer.Start()

# Afficher la fenêtre de l'interface utilisateur
$form.ShowDialog()

# Nettoyage
Close-And-Dispose-All
