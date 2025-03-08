Add-Type -AssemblyName System.Windows.Forms

# --- Definiciones para cambiar el color de la ProgressBar ---

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class Win32 {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, int lParam);
}
"@

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class ThemeHelper {
    [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
    public static extern int SetWindowTheme(IntPtr hWnd, string subAppName, string subIdList);
}
"@

$PBM_SETBARCOLOR = 0x409

# --- Depuración ---
$global:DebugMode = $true
function Debug-Print($message) {
    if ($global:DebugMode) { Write-Host "[DEBUG] $message" }
}

# --- Funciones para Dropbox ---
function Get-AccessToken {
    $AppKey = "8g8oqwp5x26h58o"
    $AppSecret = "z3690pwzqowtzjx"
    $RefreshToken = "g8E8wsPIxW8AAAAAAAAAAUVpLeEobmxg1sWlIdufgjninvxJp2x4-YLIC53n6gNe"

    $Body = @{
        refresh_token = $RefreshToken
        grant_type    = "refresh_token"
        client_id     = $AppKey
        client_secret = $AppSecret
    }
    $Headers = @{ "Content-Type" = "application/x-www-form-urlencoded" }
    $Url = "https://api.dropboxapi.com/oauth2/token"
    
    try {
        Debug-Print "Obteniendo token desde: $Url"
        $Response = Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body
        return $Response.access_token
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al renovar el token: $($_.Exception.Message)")
        return $null
    }
}

function Get-DropboxFiles($Path) {
    if ($Path -eq "/") { $Path = "" }
    $Headers = @{
        Authorization = "Bearer $global:accessToken"
        "Content-Type" = "application/json"
    }
    $Body = @{
        path               = $Path
        recursive          = $false
        include_media_info = $false
        include_deleted    = $false
    } | ConvertTo-Json -Compress
    $Url = "https://api.dropboxapi.com/2/files/list_folder"
    
    Debug-Print "Get-DropboxFiles - Path: '$Path'"
    Debug-Print "Request Body: $Body"
    
    try {
        $Response = Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body
        return $Response.entries
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al obtener archivos: $($_.ErrorDetails.Message)")
        Debug-Print "Error al obtener archivos: $($_.Exception.Message)"
        return $null
    }
}

function Get-DropboxMetadata($Path) {
    $Headers = @{
        Authorization = "Bearer $global:accessToken"
        "Content-Type" = "application/json"
    }
    $Body = @{ path = $Path } | ConvertTo-Json -Compress
    $Url = "https://api.dropboxapi.com/2/files/get_metadata"
    
    Debug-Print "Get-DropboxMetadata - Path: '$Path'"
    Debug-Print "Request Body: $Body"
    
    try {
        $Response = Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body
        return $Response
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al obtener metadata: $($_.Exception.Message)")
        Debug-Print "Error al obtener metadata: $($_.Exception.Message)"
        return $null
    }
}

# --- Descarga de archivos individuales con UI ---
function Download-DropboxFileWithProgress($FilePath, $LocalPath, $parentPanel, [bool]$ExecuteAfterDownload = $true) {
    $FileName = [System.IO.Path]::GetFileName($FilePath)
    if (-not $FileName) { return $false }
    $OutFilePath = Join-Path -Path $LocalPath -ChildPath $FileName

    # Crear panel para este archivo
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(([int]$parentPanel.Width - 10), 40)
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $fileNameLabel = New-Object System.Windows.Forms.Label
    $fileNameLabel.Text = $FileName
    $fileNameLabel.Size = New-Object System.Drawing.Size(200, 20)
    $fileNameLabel.Location = New-Object System.Drawing.Point(5, 5)
    $panel.Controls.Add($fileNameLabel)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(210, 5)
    $progressBar.Size = New-Object System.Drawing.Size(250, 20)
    $progressBar.Style = 'Continuous'
    $panel.Controls.Add($progressBar)
    $panel.Refresh()
    $null = $progressBar.Handle
    [ThemeHelper]::SetWindowTheme($progressBar.Handle, "", "")
    [Win32]::SendMessage($progressBar.Handle, $PBM_SETBARCOLOR, [IntPtr]::Zero, [System.Drawing.Color]::Red.ToArgb())

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = ""
    $statusLabel.Size = New-Object System.Drawing.Size(200, 20)
    $statusLabel.Location = New-Object System.Drawing.Point(470, 5)
    $panel.Controls.Add($statusLabel)

    $parentPanel.Controls.Add($panel)
    $parentPanel.Refresh()

    $done = New-Object System.Threading.ManualResetEvent $false
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("Authorization", "Bearer $global:accessToken")
    $dropboxAPIArg = @{ path = $FilePath } | ConvertTo-Json -Compress
    $webClient.Headers.Add("Dropbox-API-Arg", $dropboxAPIArg)

    $webClient.add_DownloadProgressChanged({
        param($sender, $args)
        $progressBar.Value = $args.ProgressPercentage
        $statusLabel.Text = "$($args.ProgressPercentage)%"
    })

    $webClient.add_DownloadFileCompleted({
        param($sender, $args)
        if ($args.Error) {
            $statusLabel.Text = "Archivo no descargado"
        } elseif ($args.Cancelled) {
            $statusLabel.Text = "Descarga cancelada"
        } else {
            $statusLabel.Text = "Completado"
        }
        $done.Set() | Out-Null
    })

    try {
        $uri = [System.Uri] "https://content.dropboxapi.com/2/files/download"
        $webClient.DownloadFileAsync($uri, $OutFilePath)
    } catch {
        $statusLabel.Text = "Error al iniciar descarga"
        $done.Set() | Out-Null
        return $false
    }

    while (-not $done.WaitOne(100)) {
        [System.Windows.Forms.Application]::DoEvents()
    }

    if (Test-Path $OutFilePath) {
        if ($ExecuteAfterDownload) {
            try {
                Start-Process -FilePath $OutFilePath
                $statusLabel.Text += " / Ejecutado"
            } catch {
                $statusLabel.Text += " / No ejecutado"
            }
        } else {
            $statusLabel.Text += " / Descargado"
        }
        return $true
    } else {
        return $false
    }
}

# --- Descarga de archivos sin UI (para uso interno en descargas de carpetas) ---
function Download-DropboxFileSilent($FilePath, $LocalPath, [bool]$ExecuteAfterDownload = $true) {
    $FileName = [System.IO.Path]::GetFileName($FilePath)
    if (-not $FileName) { return $false }
    $OutFilePath = Join-Path -Path $LocalPath -ChildPath $FileName

    $done = New-Object System.Threading.ManualResetEvent $false
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("Authorization", "Bearer $global:accessToken")
    $dropboxAPIArg = @{ path = $FilePath } | ConvertTo-Json -Compress
    $webClient.Headers.Add("Dropbox-API-Arg", $dropboxAPIArg)
    $errorOccurred = $false
    $webClient.add_DownloadFileCompleted({
        param($sender, $args)
        if ($args.Error -or $args.Cancelled) { $errorOccurred = $true }
        $done.Set() | Out-Null
    })

    try {
        $uri = [System.Uri] "https://content.dropboxapi.com/2/files/download"
        $webClient.DownloadFileAsync($uri, $OutFilePath)
    } catch {
        return $false
    }

    while (-not $done.WaitOne(100)) {
        [System.Windows.Forms.Application]::DoEvents()
    }

    if (-not (Test-Path $OutFilePath)) { return $false }
    if ($ExecuteAfterDownload) {
        try { Start-Process -FilePath $OutFilePath } catch {}
    }
    return -not $errorOccurred
}

# --- Descarga de carpeta (único progreso para toda la carpeta) ---
function Download-DropboxFolder($FolderPath, $LocalParent) {
    Debug-Print "Download-DropboxFolder - FolderPath: '$FolderPath'"
    $meta = Get-DropboxMetadata $FolderPath
    if (-not $meta -or $meta.PSObject.Properties[".tag"].Value -ne "folder") {
        Debug-Print "La ruta $FolderPath no es una carpeta válida. Abortando descarga de carpeta."
        return
    }
    
    $folderName = [System.IO.Path]::GetFileName($FolderPath)
    if (-not $folderName) { $folderName = "RootFolder" }
    $localFolderPath = Join-Path $LocalParent $folderName
    if (!(Test-Path $localFolderPath)) { New-Item -ItemType Directory -Path $localFolderPath | Out-Null }
    
    # Función interna para recopilar archivos de forma recursiva
    $allFiles = @()
    function Get-AllFiles($path) {
        $entries = Get-DropboxFiles -Path $path
        foreach ($entry in $entries) {
            if ($entry.PSObject.Properties[".tag"].Value -eq "folder") {
                Get-AllFiles $entry.path_lower
            } else {
                $allFiles += $entry.path_lower
            }
        }
    }
    Get-AllFiles $FolderPath
    $totalFiles = $allFiles.Count
    if ($totalFiles -eq 0) {
        Debug-Print "No se encontraron archivos en la carpeta."
        return
    }
    
    # Crear un panel único para la carpeta
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(680, 40)
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $folderLabel = New-Object System.Windows.Forms.Label
    $folderLabel.Text = "Descargando carpeta: $folderName"
    $folderLabel.Size = New-Object System.Drawing.Size(200, 20)
    $folderLabel.Location = New-Object System.Drawing.Point(5, 5)
    $panel.Controls.Add($folderLabel)

    $folderProgressBar = New-Object System.Windows.Forms.ProgressBar
    $folderProgressBar.Location = New-Object System.Drawing.Point(210, 5)
    $folderProgressBar.Size = New-Object System.Drawing.Size(250, 20)
    $folderProgressBar.Style = 'Continuous'
    $panel.Controls.Add($folderProgressBar)
    $panel.Refresh()
    $null = $folderProgressBar.Handle
    [ThemeHelper]::SetWindowTheme($folderProgressBar.Handle, "", "")
    [Win32]::SendMessage($folderProgressBar.Handle, $PBM_SETBARCOLOR, [IntPtr]::Zero, [System.Drawing.Color]::Red.ToArgb())

    $folderStatusLabel = New-Object System.Windows.Forms.Label
    $folderStatusLabel.Text = "0%"
    $folderStatusLabel.Size = New-Object System.Drawing.Size(200, 20)
    $folderStatusLabel.Location = New-Object System.Drawing.Point(470, 5)
    $panel.Controls.Add($folderStatusLabel)
    
    $global:downloadStatusPanel.Controls.Add($panel)
    $global:downloadStatusPanel.Refresh()

    # Descargar archivos secuencialmente sin UI individual
    $filesDownloaded = 0
    foreach ($file in $allFiles) {
         $result = Download-DropboxFileSilent $file $localFolderPath $false
         if ($result) { $filesDownloaded++ }
         $percentage = [math]::Round(($filesDownloaded / $totalFiles * 100))
         $folderProgressBar.Value = $percentage
         $folderStatusLabel.Text = "$percentage%"
    }
    if ($filesDownloaded -eq $totalFiles) {
         $folderStatusLabel.Text += " / Completado"
    } else {
         $folderStatusLabel.Text += " / Error en algunos archivos"
    }
}

# --- Interfaz gráfica ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "UTP"
$form.Size = New-Object System.Drawing.Size(700, 650)
$form.StartPosition = "CenterScreen"

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 10)
$listBox.Size = New-Object System.Drawing.Size(300, 400)
$form.Controls.Add($listBox)

$selectedFilesBox = New-Object System.Windows.Forms.ListBox
$selectedFilesBox.Location = New-Object System.Drawing.Point(350, 10)
$selectedFilesBox.Size = New-Object System.Drawing.Size(300, 350)
$form.Controls.Add($selectedFilesBox)

function Set-ButtonStyle($button) {
    $button.BackColor  = [System.Drawing.Color]::Red
    $button.ForeColor  = [System.Drawing.Color]::White
    $button.FlatStyle  = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $button.FlatAppearance.BorderSize  = 1
    $button.TextAlign  = [System.Drawing.ContentAlignment]::MiddleCenter
}

$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = "→"
$addButton.Size = New-Object System.Drawing.Size(30,30)
$addButton.Location = New-Object System.Drawing.Point(320, 200)
Set-ButtonStyle $addButton
$addButton.Add_Click({
    $selectedItem = $listBox.SelectedItem
    if ($selectedItem -and $selectedItem -ne "..") {
        $entry = $global:dropboxEntries[$selectedItem]
        if ($entry) {
            $filePath = $entry.path_lower
        } else {
            $filePath = "$global:currentPath/$selectedItem" -replace "//", "/"
        }
        $exists = $false
        foreach ($item in $selectedFilesBox.Items) {
            if ([string]$item -eq $filePath) {
                $exists = $true
                break
            }
        }
        if (-not $exists) {
            $selectedFilesBox.Items.Add($filePath)
        }
    }
})
$form.Controls.Add($addButton)

$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = "←"
$removeButton.Size = New-Object System.Drawing.Size(30,30)
$removeButton.Location = New-Object System.Drawing.Point(320, 240)
Set-ButtonStyle $removeButton
$removeButton.Add_Click({
    $selectedFilesBox.Items.Remove($selectedFilesBox.SelectedItem)
})
$form.Controls.Add($removeButton)

$moveUpButton = New-Object System.Windows.Forms.Button
$moveUpButton.Text = "↑"
$moveUpButton.Size = New-Object System.Drawing.Size(30,30)
$moveUpButton.Location = New-Object System.Drawing.Point(640, 200)
Set-ButtonStyle $moveUpButton
$moveUpButton.Add_Click({
    $index = $selectedFilesBox.SelectedIndex
    if ($index -gt 0) {
        $item = $selectedFilesBox.Items[$index]
        $selectedFilesBox.Items.RemoveAt($index)
        $selectedFilesBox.Items.Insert($index - 1, $item)
        $selectedFilesBox.SelectedIndex = $index - 1
    }
})
$form.Controls.Add($moveUpButton)

$moveDownButton = New-Object System.Windows.Forms.Button
$moveDownButton.Text = "↓"
$moveDownButton.Size = New-Object System.Drawing.Size(30,30)
$moveDownButton.Location = New-Object System.Drawing.Point(640, 240)
Set-ButtonStyle $moveDownButton
$moveDownButton.Add_Click({
    $index = $selectedFilesBox.SelectedIndex
    if ($index -lt ($selectedFilesBox.Items.Count - 1) -and $index -ge 0) {
        $item = $selectedFilesBox.Items[$index]
        $selectedFilesBox.Items.RemoveAt($index)
        $selectedFilesBox.Items.Insert($index + 1, $item)
        $selectedFilesBox.SelectedIndex = $index + 1
    }
})
$form.Controls.Add($moveDownButton)

$downloadStatusPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$downloadStatusPanel.Location = New-Object System.Drawing.Point(10, 420)
$downloadStatusPanel.Size = New-Object System.Drawing.Size(680, 180)
$downloadStatusPanel.AutoScroll = $true
$form.Controls.Add($downloadStatusPanel)
$global:downloadStatusPanel = $downloadStatusPanel

$downloadButton = New-Object System.Windows.Forms.Button
$downloadButton.Text = "Descargar"
$downloadButton.Size = New-Object System.Drawing.Size(150,30)
$downloadButton.Location = New-Object System.Drawing.Point(350, 360)
Set-ButtonStyle $downloadButton
$downloadButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $chosenPath = $folderDialog.SelectedPath
        Debug-Print "Carpeta elegida: $chosenPath"
        $global:downloadStatusPanel.Controls.Clear()
        foreach ($item in $selectedFilesBox.Items) {
            $metadata = Get-DropboxMetadata $item
            if ($metadata -and $metadata.PSObject.Properties[".tag"].Value -eq "folder") {
                Download-DropboxFolder $item $chosenPath
            } else {
                Download-DropboxFileWithProgress $item $chosenPath $global:downloadStatusPanel
            }
        }
    }
})
$form.Controls.Add($downloadButton)

$global:currentPath = ""
$global:dropboxEntries = @{}

function Update-FileList {
    if ($global:currentPath -eq "/") { $global:currentPath = "" }
    $listBox.Items.Clear()
    if ($global:currentPath -ne "") { $listBox.Items.Add("..") }
    $entries = Get-DropboxFiles -Path $global:currentPath
    if ($entries) {
        $global:dropboxEntries = @{}
        foreach ($entry in $entries) {
            if ($entry.name) {
                $listBox.Items.Add($entry.name)
                $global:dropboxEntries[$entry.name] = $entry
            }
        }
    }
}

function Get-DropboxEntry($name) {
    foreach ($key in $global:dropboxEntries.Keys) {
        if ($key.ToLower() -eq $name.ToLower()) {
            return $global:dropboxEntries[$key]
        }
    }
    return $null
}

$listBox.Add_DoubleClick({
    $selectedItem = ([string]::Concat($listBox.SelectedItem)).Trim()
    Debug-Print "Double-clicked item: '$selectedItem'"
    if (-not [string]::IsNullOrEmpty($selectedItem)) {
        if ($selectedItem -eq "..") {
            if ([string]::IsNullOrEmpty($global:currentPath) -or $global:currentPath -eq "/") {
                $global:currentPath = ""
            } else {
                $segments = $global:currentPath.Trim("/").Split("/")
                if ($segments.Length -gt 1) {
                    $global:currentPath = "/" + ($segments[0..($segments.Length - 2)] -join "/")
                } else {
                    $global:currentPath = ""
                }
            }
        } else {
            $entry = Get-DropboxEntry $selectedItem
            if ($entry -and $entry.PSObject.Properties[".tag"].Value -eq "folder") {
                if ([string]::IsNullOrEmpty($global:currentPath) -or $global:currentPath -eq "/") {
                    $global:currentPath = "/" + $selectedItem
                } else {
                    $global:currentPath = $global:currentPath + "/" + $selectedItem
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("El elemento seleccionado no es una carpeta.")
            }
        }
        Debug-Print "New currentPath: '$global:currentPath'"
        Update-FileList
    }
})

$global:accessToken = Get-AccessToken
if ($global:accessToken) {
    Update-FileList
    $form.ShowDialog()
}
