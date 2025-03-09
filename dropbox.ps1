Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

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
        Debug-Print "Error al obtener metadata: $($_.Exception.Message)"
        return $null
    }
}

function Upload-DropboxFile($LocalFilePath, $DropboxDestinationFolder) {
    $FileName = [System.IO.Path]::GetFileName($LocalFilePath)
    if ([string]::IsNullOrEmpty($DropboxDestinationFolder) -or $DropboxDestinationFolder -eq "/") {
        $DropboxPath = "/$FileName"
    } else {
        $DropboxPath = "$DropboxDestinationFolder/$FileName"
    }
    $Url = "https://content.dropboxapi.com/2/files/upload"
    $Headers = @{
         "Authorization" = "Bearer $global:accessToken"
         "Dropbox-API-Arg" = (ConvertTo-Json -Compress @{
            path = $DropboxPath
            mode = "add"
            autorename = $true
            mute = $false
            strict_conflict = $false
         })
         "Content-Type" = "application/octet-stream"
    }
    try {
         Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -InFile $LocalFilePath -ContentType "application/octet-stream" | Out-Null
         return $true
    } catch {
         Debug-Print "Error al subir archivo: $($_.Exception.Message)"
         return $false
    }
}

function Upload-DropboxFolder($LocalFolderPath, $DropboxDestinationFolder) {
    $FolderName = [System.IO.Path]::GetFileName($LocalFolderPath)
    if ([string]::IsNullOrEmpty($DropboxDestinationFolder) -or $DropboxDestinationFolder -eq "/") {
        $DropboxFolderPath = "/$FolderName"
    } else {
        $DropboxFolderPath = "$DropboxDestinationFolder/$FolderName"
    }
    # Crear carpeta en Dropbox
    $Url = "https://api.dropboxapi.com/2/files/create_folder_v2"
    $Headers = @{
         "Authorization" = "Bearer $global:accessToken"
         "Content-Type"  = "application/json"
    }
    $Body = @{
         path = $DropboxFolderPath
         autorename = $true
    } | ConvertTo-Json -Compress
    try {
         Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body | Out-Null
    } catch {
         Debug-Print "Error al crear carpeta (posiblemente ya exista): $($_.Exception.Message)"
    }
    # Subir archivos inmediatos
    Get-ChildItem -Path $LocalFolderPath -File | ForEach-Object {
         Upload-DropboxFile $_.FullName $DropboxFolderPath | Out-Null
    }
    # Procesar subcarpetas de forma recursiva
    Get-ChildItem -Path $LocalFolderPath -Directory | ForEach-Object {
         Upload-DropboxFolder $_.FullName $DropboxFolderPath
    }
}

function Copy-DropboxItem($fromPath, $toPath) {
    $Headers = @{
        Authorization = "Bearer $global:accessToken"
        "Content-Type" = "application/json"
    }
    $Body = @{
        from_path = $fromPath
        to_path   = $toPath
        autorename = $true
    } | ConvertTo-Json -Compress
    $Url = "https://api.dropboxapi.com/2/files/copy_v2"
    try {
        Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body | Out-Null
        return $true
    } catch {
        Debug-Print "Error al copiar: $($_.Exception.Message)"
        return $false
    }
}

function Move-DropboxItem($fromPath, $toPath) {
    $Headers = @{
        Authorization = "Bearer $global:accessToken"
        "Content-Type" = "application/json"
    }
    $Body = @{
        from_path = $fromPath
        to_path   = $toPath
        autorename = $true
    } | ConvertTo-Json -Compress
    $Url = "https://api.dropboxapi.com/2/files/move_v2"
    try {
        Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body | Out-Null
        return $true
    } catch {
        Debug-Print "Error al mover: $($_.Exception.Message)"
        return $false
    }
}

function Delete-DropboxItem($path) {
    $Headers = @{
        Authorization = "Bearer $global:accessToken"
        "Content-Type" = "application/json"
    }
    $Body = @{ path = $path } | ConvertTo-Json -Compress
    $Url = "https://api.dropboxapi.com/2/files/delete_v2"
    try {
        Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body | Out-Null
        return $true
    } catch {
        Debug-Print "Error al eliminar: $($_.Exception.Message)"
        return $false
    }
}

function Download-DropboxFileWithProgress($FilePath, $LocalPath, $parentPanel, [bool]$ExecuteAfterDownload = $true) {
    $FileName = [System.IO.Path]::GetFileName($FilePath)
    if (-not $FileName) { return $false }
    $OutFilePath = Join-Path -Path $LocalPath -ChildPath $FileName

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

$fileListView = New-Object System.Windows.Forms.ListView
$fileListView.Location = New-Object System.Drawing.Point(10,10)
$fileListView.Size = New-Object System.Drawing.Size(300,400)
$fileListView.View = [System.Windows.Forms.View]::List

$imageList = New-Object System.Windows.Forms.ImageList
$folderIcon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\explorer.exe").ToBitmap()
$fileIcon = [System.Drawing.SystemIcons]::Information.ToBitmap()
$imageList.Images.Add("folder", $folderIcon)
$imageList.Images.Add("file", $fileIcon)

$fileListView.SmallImageList = $imageList
$form.Controls.Add($fileListView)

$selectedFilesBox = New-Object System.Windows.Forms.ListBox
$selectedFilesBox.Location = New-Object System.Drawing.Point(350,10)
$selectedFilesBox.Size = New-Object System.Drawing.Size(300,350)
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
$addButton.Location = New-Object System.Drawing.Point(320,200)
Set-ButtonStyle $addButton
$addButton.Add_Click({
    if ($fileListView.SelectedItems.Count -gt 0) {
         $selectedItem = $fileListView.SelectedItems[0].Text
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
    }
})
$form.Controls.Add($addButton)

$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = "←"
$removeButton.Size = New-Object System.Drawing.Size(30,30)
$removeButton.Location = New-Object System.Drawing.Point(320,240)
Set-ButtonStyle $removeButton
$removeButton.Add_Click({
    $selectedFilesBox.Items.Remove($selectedFilesBox.SelectedItem)
})
$form.Controls.Add($removeButton)

$moveUpButton = New-Object System.Windows.Forms.Button
$moveUpButton.Text = "↑"
$moveUpButton.Size = New-Object System.Drawing.Size(30,30)
$moveUpButton.Location = New-Object System.Drawing.Point(640,200)
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
$moveDownButton.Location = New-Object System.Drawing.Point(640,240)
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
$downloadStatusPanel.Location = New-Object System.Drawing.Point(10,420)
$downloadStatusPanel.Size = New-Object System.Drawing.Size(680,180)
$downloadStatusPanel.AutoScroll = $true
$form.Controls.Add($downloadStatusPanel)
$global:downloadStatusPanel = $downloadStatusPanel

$downloadButton = New-Object System.Windows.Forms.Button
$downloadButton.Text = "Descargar"
$downloadButton.Size = New-Object System.Drawing.Size(150,30)
$downloadButton.Location = New-Object System.Drawing.Point(350,360)
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

# --- Botón para subir archivos y carpetas (reemplazando el MessageBox) ---
$uploadButton = New-Object System.Windows.Forms.Button
$uploadButton.Text = "Subir"
$uploadButton.Size = New-Object System.Drawing.Size(150,30)
$uploadButton.Location = New-Object System.Drawing.Point(350,400)
Set-ButtonStyle $uploadButton
$uploadButton.Add_Click({

    # Creamos un formulario pequeño para elegir si subir Archivo o Carpeta
    $uploadForm = New-Object System.Windows.Forms.Form
    $uploadForm.Text = "Seleccionar tipo de subida"
    $uploadForm.Size = New-Object System.Drawing.Size(300,150)
    $uploadForm.StartPosition = "CenterParent"
    $uploadForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $uploadForm.MaximizeBox = $false
    $uploadForm.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "¿Qué desea subir?"
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point(10,10)
    $uploadForm.Controls.Add($lbl)

    $btnArchivo = New-Object System.Windows.Forms.Button
    $btnArchivo.Text = "Archivo"
    $btnArchivo.Size = New-Object System.Drawing.Size(70,30)
    $btnArchivo.Location = New-Object System.Drawing.Point(10,50)
    $btnArchivo.Add_Click({
        $uploadForm.Tag = "file"
        $uploadForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $uploadForm.Close()
    })
    $uploadForm.Controls.Add($btnArchivo)

    $btnCarpeta = New-Object System.Windows.Forms.Button
    $btnCarpeta.Text = "Carpeta"
    $btnCarpeta.Size = New-Object System.Drawing.Size(70,30)
    $btnCarpeta.Location = New-Object System.Drawing.Point(90,50)
    $btnCarpeta.Add_Click({
        $uploadForm.Tag = "folder"
        $uploadForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $uploadForm.Close()
    })
    $uploadForm.Controls.Add($btnCarpeta)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Size = New-Object System.Drawing.Size(70,30)
    $btnCancelar.Location = New-Object System.Drawing.Point(170,50)
    $btnCancelar.Add_Click({
        $uploadForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $uploadForm.Close()
    })
    $uploadForm.Controls.Add($btnCancelar)

    $dialogResult = $uploadForm.ShowDialog($form)

    # Si el usuario eligió Archivo o Carpeta:
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $option = $uploadForm.Tag
        if ($option -eq "file") {
            $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $fileDialog.Multiselect = $true
            if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                foreach ($file in $fileDialog.FileNames) {
                    Upload-DropboxFile $file $global:currentPath | Out-Null
                }
                Update-FileList
            }
        } elseif ($option -eq "folder") {
            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                Upload-DropboxFolder $folderDialog.SelectedPath $global:currentPath
                Update-FileList
            }
        }
    }
})
$form.Controls.Add($uploadButton)

# --- Botón para crear carpetas ---
$createFolderButton = New-Object System.Windows.Forms.Button
$createFolderButton.Text = "Crear carpeta"
$createFolderButton.Size = New-Object System.Drawing.Size(150,30)
$createFolderButton.Location = New-Object System.Drawing.Point(510,400)
Set-ButtonStyle $createFolderButton
$createFolderButton.Add_Click({
    $folderName = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese el nombre de la nueva carpeta", "Crear carpeta", "Nueva carpeta")
    if ($folderName -ne "") {
         if ([string]::IsNullOrEmpty($global:currentPath) -or $global:currentPath -eq "/") {
             $DropboxFolderPath = "/$folderName"
         } else {
             $DropboxFolderPath = "$global:currentPath/$folderName"
         }
         $Url = "https://api.dropboxapi.com/2/files/create_folder_v2"
         $Headers = @{
             Authorization = "Bearer $global:accessToken"
             "Content-Type"  = "application/json"
         }
         $Body = @{
             path = $DropboxFolderPath
             autorename = $true
         } | ConvertTo-Json -Compress
         try {
             Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body | Out-Null
             Update-FileList
         } catch {
             Debug-Print "Error al crear carpeta: $($_.Exception.Message)"
         }
    }
})
$form.Controls.Add($createFolderButton)

$global:clipboardItem = $null
$global:clipboardOperation = $null

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuCopy = New-Object System.Windows.Forms.ToolStripMenuItem("Copiar")
$menuCut = New-Object System.Windows.Forms.ToolStripMenuItem("Cortar")
$menuPaste = New-Object System.Windows.Forms.ToolStripMenuItem("Pegar")
$menuDelete = New-Object System.Windows.Forms.ToolStripMenuItem("Eliminar")
$contextMenu.Items.Add($menuCopy)
$contextMenu.Items.Add($menuCut)
$contextMenu.Items.Add($menuPaste)
$contextMenu.Items.Add($menuDelete)
$fileListView.ContextMenuStrip = $contextMenu

$menuCopy.Add_Click({
    if ($fileListView.SelectedItems.Count -gt 0) {
         $selectedItem = $fileListView.SelectedItems[0].Text
         if ($selectedItem -and $selectedItem -ne "..") {
             $entry = $global:dropboxEntries[$selectedItem]
             if ($entry) {
                 $global:clipboardItem = $entry.path_lower
             } else {
                 $global:clipboardItem = "$global:currentPath/$selectedItem" -replace "//", "/"
             }
             $global:clipboardOperation = "copy"
         }
    }
})

$menuCut.Add_Click({
    if ($fileListView.SelectedItems.Count -gt 0) {
         $selectedItem = $fileListView.SelectedItems[0].Text
         if ($selectedItem -and $selectedItem -ne "..") {
             $entry = $global:dropboxEntries[$selectedItem]
             if ($entry) {
                 $global:clipboardItem = $entry.path_lower
             } else {
                 $global:clipboardItem = "$global:currentPath/$selectedItem" -replace "//", "/"
             }
             $global:clipboardOperation = "cut"
         }
    }
})

$menuPaste.Add_Click({
    if ($global:clipboardItem) {
         $destination = if ([string]::IsNullOrEmpty($global:currentPath) -or $global:currentPath -eq "/") {
             "/$([System.IO.Path]::GetFileName($global:clipboardItem))"
         } else {
             "$global:currentPath/$([System.IO.Path]::GetFileName($global:clipboardItem))"
         }
         if ($global:clipboardOperation -eq "copy") {
             $result = Copy-DropboxItem $global:clipboardItem $destination
         } elseif ($global:clipboardOperation -eq "cut") {
             $result = Move-DropboxItem $global:clipboardItem $destination
         }
         if ($result) {
             Update-FileList
             $global:clipboardItem = $null
             $global:clipboardOperation = $null
         }
    }
})

$menuDelete.Add_Click({
    if ($fileListView.SelectedItems.Count -gt 0) {
         $selectedItem = $fileListView.SelectedItems[0].Text
         if ($selectedItem -and $selectedItem -ne "..") {
             $entry = $global:dropboxEntries[$selectedItem]
             if ($entry) {
                 $pathToDelete = $entry.path_lower
             } else {
                 $pathToDelete = "$global:currentPath/$selectedItem" -replace "//", "/"
             }
             if (Delete-DropboxItem $pathToDelete) {
                 Update-FileList
             }
         }
    }
})

$global:currentPath = ""
$global:dropboxEntries = @{}

function Update-FileList {
    if ($global:currentPath -eq "/") { $global:currentPath = "" }
    $fileListView.Items.Clear()
    if ($global:currentPath -ne "") {
         $upItem = New-Object System.Windows.Forms.ListViewItem("..")
         $upItem.ImageKey = "folder"
         $fileListView.Items.Add($upItem)
    }
    $entries = Get-DropboxFiles -Path $global:currentPath
    if ($entries) {
        $global:dropboxEntries = @{}
        foreach ($entry in $entries) {
            if ($entry.name) {
                $lvi = New-Object System.Windows.Forms.ListViewItem($entry.name)
                if ($entry.PSObject.Properties[".tag"].Value -eq "folder") {
                    $lvi.ImageKey = "folder"
                } else {
                    $lvi.ImageKey = "file"
                }
                $fileListView.Items.Add($lvi)
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

$fileListView.Add_DoubleClick({
    if ($fileListView.SelectedItems.Count -gt 0) {
         $selectedItem = $fileListView.SelectedItems[0].Text
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
                 }
             }
             Debug-Print "New currentPath: '$global:currentPath'"
             Update-FileList
         }
    }
})

$global:accessToken = Get-AccessToken
if ($global:accessToken) {
    Update-FileList
    $form.ShowDialog()
}
