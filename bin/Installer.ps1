Add-Type -AssemblyName PresentationFramework

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

function Show-InstallerUI {
    [System.Windows.MessageBoxButton]$buttons = "YesNo"
    [System.Windows.MessageBoxImage]$icon = "Question"
    
    $message = "Bienvenue dans l'installateur de Mission Flash." + [Environment]::NewLine +
               "Cette installation va installer les dépendances Python nécessaires, copier les fichiers dans Documents et créer un raccourci sur le bureau." + [Environment]::NewLine +
               "Acceptez-vous de poursuivre l'installation ?"
               
    $caption = "Installateur Mission Flash"
    
    $result = [System.Windows.MessageBox]::Show($message, $caption, $buttons, $icon)
    
    if ($result -eq "Yes") {
        # Création fenêtre de progression
        $window = New-Object System.Windows.Window
        $window.Title = "Installation en cours..."
        $window.Width = 400
        $window.Height = 150
        $window.WindowStartupLocation = "CenterScreen"
        $window.ResizeMode = "NoResize"
        
        $stackPanel = New-Object System.Windows.Controls.StackPanel
        $stackPanel.Orientation = "Vertical"
        $stackPanel.Margin = "10"
        
        $textBlock = New-Object System.Windows.Controls.TextBlock
        $textBlock.Text = "Installation en cours, veuillez patienter..."
        $textBlock.Margin = "0,0,0,10"
        $textBlock.FontSize = 14
        $stackPanel.Children.Add($textBlock)
        
        $progressBar = New-Object System.Windows.Controls.ProgressBar
        $progressBar.Minimum = 0
        $progressBar.Maximum = 100
        $progressBar.Height = 20
        $progressBar.IsIndeterminate = $true
        $stackPanel.Children.Add($progressBar)
        
        $window.Content = $stackPanel
        $window.Show()

        # Installer les dépendances Python (en job pour ne pas bloquer l'UI)
        Start-Job -ScriptBlock {
            python -m pip install --upgrade pip
            python -m pip install pyautogui keyboard pygetwindow
        } | Wait-Job
        
        # Copier les fichiers dans Documents\MissionFlash
        $destFolder = Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "MissionFlash"
        if (-Not (Test-Path $destFolder)) {
            New-Item -Path $destFolder -ItemType Directory | Out-Null
        }
        
        # Copier tous les fichiers du dossier actuel vers $destFolder
        $sourceFolder = $PSScriptRoot
        Get-ChildItem -Path $sourceFolder -File | ForEach-Object {
            Copy-Item -Path $_.FullName -Destination $destFolder -Force
        }

        # Création raccourci sur le bureau
        $shell = New-Object -ComObject WScript.Shell
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $shortcutPath = Join-Path -Path $desktopPath -ChildPath "Mission Flash.lnk"

        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "pythonw.exe"
        $shortcut.Arguments = "`"$destFolder\main.pyw`""
        $shortcut.WorkingDirectory = $destFolder
        $shortcut.WindowStyle = 1
        $shortcut.Description = "Lancer Mission Flash"
        $shortcut.IconLocation = "$destFolder\mission-flash.ico, 0"
        $shortcut.Save()
        
        # Mise à jour interface fin
        $progressBar.IsIndeterminate = $false
        $progressBar.Value = 100
        $textBlock.Text = "Installation terminée. Vous pouvez accéder à l'application depuis le bureau."
        
        Start-Sleep -Seconds 5
        $window.Close()
        
        exit 0
    } else {
        [System.Windows.MessageBox]::Show("Installation annulée.", "Installateur Mission Flash", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        exit 1
    }
}

Show-InstallerUI
