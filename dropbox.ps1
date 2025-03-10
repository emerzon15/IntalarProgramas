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

# --- Añadimos funciones para obtener el ícono asociado a una extensión ---
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class ShellIcon {
    [StructLayout(LayoutKind.Sequential)]
    public struct SHFILEINFO {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    };

    public const uint SHGFI_ICON = 0x100;
    public const uint SHGFI_SMALLICON = 0x1;

    [DllImport("shell32.dll")]
    public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes,
        ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);
}
"@

function Get-FileIcon {
    param([string]$filePath)
    $shinfo = New-Object ShellIcon+SHFILEINFO
    $size = [System.Runtime.InteropServices.Marshal]::SizeOf($shinfo)
    $flags = [ShellIcon]::SHGFI_ICON -bor [ShellIcon]::SHGFI_SMALLICON
    [ShellIcon]::SHGetFileInfo($filePath, 0, [ref]$shinfo, $size, $flags) | Out-Null
    $icon = [System.Drawing.Icon]::FromHandle($shinfo.hIcon)
    $bitmap = $icon.ToBitmap()
    [System.Runtime.InteropServices.Marshal]::DestroyIcon($shinfo.hIcon) | Out-Null
    return $bitmap
}

# --- Depuración ---
$global:DebugMode = $true
function Debug-Print($message) {
    if ($global:DebugMode) { Write-Host "[DEBUG] $message" }
}

# ====================================
# == Funciones de Dropbox (API) ==
# ====================================

function Get-AccessToken {
    # Reemplaza con tus credenciales
    $AppKey       = "8g8oqwp5x26h58o"
    $AppSecret    = "z3690pwzqowtzjx"
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
        Debug-Print "Error al obtener token: $($_.Exception.Message)"
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
    $allEntries = @()
    try {
        $Response = Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body
        $allEntries += $Response.entries
        while ($Response.has_more -eq $true) {
            $cursor = $Response.cursor
            $UrlContinue = "https://api.dropboxapi.com/2/files/list_folder/continue"
            $BodyContinue = @{ cursor = $cursor } | ConvertTo-Json -Compress
            $Response = Invoke-RestMethod -Uri $UrlContinue -Method Post -Headers $Headers -Body $BodyContinue
            $allEntries += $Response.entries
        }
        return $allEntries
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
    $Url = "https://api.dropboxapi.com/2/files/create_folder_v2"
    $Headers = @{
        "Authorization" = "Bearer $global:accessToken"
        "Content-Type"  = "application/json"
    }
    $Body = @{
        path       = $DropboxFolderPath
        autorename = $true
    } | ConvertTo-Json -Compress
    try {
        Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body | Out-Null
    } catch {
        Debug-Print "Error al crear carpeta: $($_.Exception.Message)"
    }
    Get-ChildItem -Path $LocalFolderPath -File | ForEach-Object {
        Upload-DropboxFile $_.FullName $DropboxFolderPath | Out-Null
    }
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
        from_path  = $fromPath
        to_path    = $toPath
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
        from_path  = $fromPath
        to_path    = $toPath
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

# ========================================
# == Funciones de Descarga (Sin Execute) ==
# ========================================

function Download-DropboxFileWithProgressNoExecute($FilePath, $LocalPath, $parentPanel) {
    $FileName = [System.IO.Path]::GetFileName($FilePath)
    if (-not $FileName) {
        Debug-Print "No se pudo determinar nombre para '$FilePath'"
        return $null
    }
    $OutFilePath = Join-Path -Path $LocalPath -ChildPath $FileName
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(([int]$parentPanel.Width - 25), 40)
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $fileNameLabel = New-Object System.Windows.Forms.Label
    $fileNameLabel.Text = $FileName
    $fileNameLabel.Size = New-Object System.Drawing.Size(200,20)
    $fileNameLabel.Location = New-Object System.Drawing.Point(5,5)
    $panel.Controls.Add($fileNameLabel)
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(210,5)
    $progressBar.Size = New-Object System.Drawing.Size(250,20)
    $progressBar.Style = 'Continuous'
    $panel.Controls.Add($progressBar)
    $panel.Refresh()
    $null = $progressBar.Handle
    [ThemeHelper]::SetWindowTheme($progressBar.Handle, "", "")
    [Win32]::SendMessage($progressBar.Handle, $PBM_SETBARCOLOR, [IntPtr]::Zero, [System.Drawing.Color]::Red.ToArgb())
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = ""
    $statusLabel.Size = New-Object System.Drawing.Size(150,20)
    $statusLabel.Location = New-Object System.Drawing.Point(470,5)
    $panel.Controls.Add($statusLabel)
    $parentPanel.Controls.Add($panel)
    $parentPanel.ScrollControlIntoView($panel)
    $parentPanel.Refresh()
    $done = New-Object System.Threading.ManualResetEvent $false
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("Authorization", "Bearer $global:accessToken")
    $dropboxAPIArg = @{ path = $FilePath } | ConvertTo-Json -Compress
    $webClient.Headers.Add("Dropbox-API-Arg", $dropboxAPIArg)
    $webClient.add_DownloadProgressChanged({
        param($sender, $args)
        $progressBar.Value = $args.ProgressPercentage
        $statusLabel.Text  = "$($args.ProgressPercentage)%"
    })
    $webClient.add_DownloadFileCompleted({
        param($sender, $args)
        if ($args.Error) {
            $statusLabel.Text = "Archivo no descargado"
            Debug-Print "Error: $($args.Error.Message)"
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
        Debug-Print "Error al iniciar descarga: $($_.Exception.Message)"
        $done.Set() | Out-Null
        return $null
    }
    while (-not $done.WaitOne(100)) {
        [System.Windows.Forms.Application]::DoEvents()
    }
    if (Test-Path $OutFilePath) {
        return $OutFilePath
    } else {
        return $null
    }
}

function Download-DropboxFileSilent($FilePath, $LocalPath) {
    $FileName = [System.IO.Path]::GetFileName($FilePath)
    if (-not $FileName) { return $null }
    $OutFilePath = Join-Path -Path $LocalPath -ChildPath $FileName
    $done = New-Object System.Threading.ManualResetEvent $false
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("Authorization", "Bearer $global:accessToken")
    $dropboxAPIArg = @{ path = $FilePath } | ConvertTo-Json -Compress
    $webClient.Headers.Add("Dropbox-API-Arg", $dropboxAPIArg)
    $errorOccurred = $false
    $webClient.add_DownloadFileCompleted({
        param($sender, $args)
        if ($args.Error -or $args.Cancelled) {
            $errorOccurred = $true
            Debug-Print "Error/Cancel en descarga silent: $($args.Error)"
        }
        $done.Set() | Out-Null
    })
    try {
        $uri = [System.Uri] "https://content.dropboxapi.com/2/files/download"
        $webClient.DownloadFileAsync($uri, $OutFilePath)
    } catch {
        Debug-Print "Error al iniciar descarga silent: $($_.Exception.Message)"
        return $null
    }
    while (-not $done.WaitOne(100)) {
        [System.Windows.Forms.Application]::DoEvents()
    }
    if (-not (Test-Path $OutFilePath) -or $errorOccurred) {
        return $null
    }
    return $OutFilePath
}

function Download-DropboxFolderNoExecute($FolderPath, $LocalParent, $parentPanel) {
    $resultPaths = @()
    $meta = Get-DropboxMetadata $FolderPath
    if (-not $meta -or $meta.PSObject.Properties[".tag"].Value -ne "folder") {
        Debug-Print "La ruta $FolderPath no es carpeta. Abortando."
        return $resultPaths
    }
    $folderName = [System.IO.Path]::GetFileName($FolderPath)
    if (-not $folderName) { $folderName = "RootFolder" }
    $localFolderPath = Join-Path $LocalParent $folderName
    if (!(Test-Path $localFolderPath)) {
        New-Item -ItemType Directory -Path $localFolderPath | Out-Null
    }
    # Usamos un scriptblock para recursión y obtener TODOS los archivos
    $allFiles = @()
    $GetAllFiles = {
        param($path)
        $entries = Get-DropboxFiles -Path $path
        foreach ($entry in $entries) {
            if ($entry.PSObject.Properties[".tag"].Value -eq "folder") {
                & $GetAllFiles $entry.path_lower
            } else {
                $allFiles += $entry.path_lower
            }
        }
    }
    & $GetAllFiles $FolderPath
    $totalFiles = $allFiles.Count
    if ($totalFiles -eq 0) {
        Debug-Print "No hay archivos en la carpeta $FolderPath"
        return $resultPaths
    }
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(([int]$parentPanel.Width - 25), 40)
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $folderLabel = New-Object System.Windows.Forms.Label
    $folderLabel.Text = "Descargando carpeta: $folderName"
    $folderLabel.Size = New-Object System.Drawing.Size(200,20)
    $folderLabel.Location = New-Object System.Drawing.Point(5,5)
    $panel.Controls.Add($folderLabel)
    $folderProgressBar = New-Object System.Windows.Forms.ProgressBar
    $folderProgressBar.Location = New-Object System.Drawing.Point(210,5)
    $folderProgressBar.Size = New-Object System.Drawing.Size(250,20)
    $folderProgressBar.Style = 'Continuous'
    $panel.Controls.Add($folderProgressBar)
    $panel.Refresh()
    $null = $folderProgressBar.Handle
    [ThemeHelper]::SetWindowTheme($folderProgressBar.Handle, "", "")
    [Win32]::SendMessage($folderProgressBar.Handle, $PBM_SETBARCOLOR, [IntPtr]::Zero, [System.Drawing.Color]::Red.ToArgb())
    $folderStatusLabel = New-Object System.Windows.Forms.Label
    $folderStatusLabel.Text = "0%"
    $folderStatusLabel.Size = New-Object System.Drawing.Size(150,20)
    $folderStatusLabel.Location = New-Object System.Drawing.Point(470,5)
    $panel.Controls.Add($folderStatusLabel)
    $parentPanel.Controls.Add($panel)
    $parentPanel.ScrollControlIntoView($panel)
    $parentPanel.Refresh()
    $filesDownloaded = 0
    foreach ($file in $allFiles) {
        $localPath = Download-DropboxFileSilent $file $localFolderPath
        if ($localPath) {
            $filesDownloaded++
            $resultPaths += $localPath
        }
        $percentage = [math]::Round(($filesDownloaded / $totalFiles * 100))
        $folderProgressBar.Value = $percentage
        $folderStatusLabel.Text  = "$percentage%"
    }
    if ($filesDownloaded -eq $totalFiles) {
        $folderStatusLabel.Text += " / Completado"
    } else {
        $folderStatusLabel.Text += " / Error en algunos archivos"
    }
    return $resultPaths
}

# ============================================
# == Interfaz Gráfica (Código Completo) ==
# ============================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "UTP"
$form.Size = New-Object System.Drawing.Size(900,600)
$form.StartPosition = "CenterScreen"

# Creamos el ImageList y agregamos el ícono para carpetas
$imageList = New-Object System.Windows.Forms.ImageList
$folderIcon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\explorer.exe").ToBitmap()
$imageList.Images.Add("folder", $folderIcon)
# Opcional: ícono por defecto para archivos si no se encuentra según extensión
if (-not $imageList.Images.ContainsKey("file")) {
    $defaultFileIcon = [System.Drawing.SystemIcons]::Information.ToBitmap()
    $imageList.Images.Add("file", $defaultFileIcon)
}

# ListView (izquierda) para mostrar archivos/carpetas de Dropbox
$fileListView = New-Object System.Windows.Forms.ListView
$fileListView.Location = New-Object System.Drawing.Point(10,10)
$fileListView.Size = New-Object System.Drawing.Size(350,400)
$fileListView.View = [System.Windows.Forms.View]::List
$fileListView.SmallImageList = $imageList
$form.Controls.Add($fileListView)

# ListBox (derecha) para los ítems seleccionados para descarga
$selectedFilesBox = New-Object System.Windows.Forms.ListBox
$selectedFilesBox.Location = New-Object System.Drawing.Point(460,10)
$selectedFilesBox.Size = New-Object System.Drawing.Size(350,400)
$form.Controls.Add($selectedFilesBox)

function Set-ButtonStyle($button) {
    $button.BackColor  = [System.Drawing.Color]::Red
    $button.ForeColor  = [System.Drawing.Color]::White
    $button.FlatStyle  = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $button.FlatAppearance.BorderSize  = 1
    $button.TextAlign  = [System.Drawing.ContentAlignment]::MiddleCenter
}

# Botón para pasar ítems de la izquierda a la derecha
$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = "→"
$addButton.Size = New-Object System.Drawing.Size(30,30)
$addButton.Location = New-Object System.Drawing.Point(420,150)
Set-ButtonStyle $addButton
$addButton.Add_Click({
    if ($fileListView.SelectedItems.Count -gt 0) {
        $selectedItem = $fileListView.SelectedItems[0].Text
        if ($selectedItem -and $selectedItem -ne "..") {
            $entry = $global:dropboxEntries[$selectedItem]
            $filePath = if ($entry) { $entry.path_lower } else { "$global:currentPath/$selectedItem" -replace "//","/" }
            if (-not $selectedFilesBox.Items.Contains($filePath)) {
                $selectedFilesBox.Items.Add($filePath)
            }
        }
    }
})
$form.Controls.Add($addButton)

$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = "←"
$removeButton.Size = New-Object System.Drawing.Size(30,30)
$removeButton.Location = New-Object System.Drawing.Point(420,190)
Set-ButtonStyle $removeButton
$removeButton.Add_Click({
    $selectedFilesBox.Items.Remove($selectedFilesBox.SelectedItem)
})
$form.Controls.Add($removeButton)

# Botones para reordenar la lista de archivos seleccionados (flecha arriba y flecha abajo)
$moveUpButton = New-Object System.Windows.Forms.Button
$moveUpButton.Text = "↑"
$moveUpButton.Size = New-Object System.Drawing.Size(30,30)
$moveUpButton.Location = New-Object System.Drawing.Point(820,150)
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
$moveDownButton.Location = New-Object System.Drawing.Point(820,190)
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

# Panel para mostrar el progreso de descarga
$downloadStatusPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$downloadStatusPanel.Location = New-Object System.Drawing.Point(10,500)
$downloadStatusPanel.Size = New-Object System.Drawing.Size(870,60)
$downloadStatusPanel.AutoScroll = $true
$form.Controls.Add($downloadStatusPanel)
$global:downloadStatusPanel = $downloadStatusPanel

# Botón "Subir"
$uploadButton = New-Object System.Windows.Forms.Button
$uploadButton.Text = "Subir"
$uploadButton.Size = New-Object System.Drawing.Size(100,40)
$uploadButton.Location = New-Object System.Drawing.Point(10,420)
Set-ButtonStyle $uploadButton
$uploadButton.Add_Click({
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

# Botón "Crear carpeta"
$createFolderButton = New-Object System.Windows.Forms.Button
$createFolderButton.Text = "Crear carpeta"
$createFolderButton.Size = New-Object System.Drawing.Size(100,40)
$createFolderButton.Location = New-Object System.Drawing.Point(120,420)
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
            path       = $DropboxFolderPath
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

# Botón "Descargar": primero descarga todos los ítems y luego los ejecuta en orden.
$downloadButton = New-Object System.Windows.Forms.Button
$downloadButton.Text = "Descargar"
$downloadButton.Size = New-Object System.Drawing.Size(100,40)
$downloadButton.Location = New-Object System.Drawing.Point(230,420)
Set-ButtonStyle $downloadButton
$downloadButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $chosenPath = $folderDialog.SelectedPath
        Debug-Print "Carpeta elegida para descargar: $chosenPath"
        $global:downloadStatusPanel.Controls.Clear()
        $allDownloadedPaths = @()
        foreach ($item in $selectedFilesBox.Items) {
            $metadata = Get-DropboxMetadata $item
            if ($metadata -and $metadata.PSObject.Properties[".tag"].Value -eq "folder") {
                $paths = Download-DropboxFolderNoExecute $item $chosenPath $global:downloadStatusPanel
                $allDownloadedPaths += $paths
            } else {
                $localPath = Download-DropboxFileWithProgressNoExecute $item $chosenPath $global:downloadStatusPanel
                if ($localPath) { $allDownloadedPaths += $localPath }
            }
        }
        foreach ($localFile in $allDownloadedPaths) {
            Debug-Print "Ejecutando: $localFile"
            try {
                Start-Process -FilePath $localFile
            } catch {
                Debug-Print "No se pudo ejecutar $localFile $($_.Exception.Message)"
            }
        }
    }
})
$form.Controls.Add($downloadButton)

# Menú contextual para el ListView (Copiar, Cortar, Pegar, Eliminar)
$global:clipboardItem = $null
$global:clipboardOperation = $null
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuCopy   = New-Object System.Windows.Forms.ToolStripMenuItem("Copiar")
$menuCut    = New-Object System.Windows.Forms.ToolStripMenuItem("Cortar")
$menuPaste  = New-Object System.Windows.Forms.ToolStripMenuItem("Pegar")
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
            $global:clipboardItem = if ($entry) { $entry.path_lower } else { "$global:currentPath/$selectedItem" -replace "//","/" }
            $global:clipboardOperation = "copy"
        }
    }
})
$menuCut.Add_Click({
    if ($fileListView.SelectedItems.Count -gt 0) {
        $selectedItem = $fileListView.SelectedItems[0].Text
        if ($selectedItem -and $selectedItem -ne "..") {
            $entry = $global:dropboxEntries[$selectedItem]
            $global:clipboardItem = if ($entry) { $entry.path_lower } else { "$global:currentPath/$selectedItem" -replace "//","/" }
            $global:clipboardOperation = "cut"
        }
    }
})
$menuPaste.Add_Click({
    if ($global:clipboardItem) {
        $destFile = [System.IO.Path]::GetFileName($global:clipboardItem)
        $destination = if ([string]::IsNullOrEmpty($global:currentPath) -or $global:currentPath -eq "/") {
            "/$destFile"
        } else {
            "$global:currentPath/$destFile"
        }
        $result = $false
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
            $pathToDelete = if ($entry) { $entry.path_lower } else { "$global:currentPath/$selectedItem" -replace "//","/" }
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
    # Se obtiene y ordena la lista de archivos y carpetas:
    # Primero se colocan las carpetas (valor 0) y luego los archivos (valor 1), ambos ordenados alfabéticamente.
    $entries = Get-DropboxFiles -Path $global:currentPath | Sort-Object -Property { if($_.PSObject.Properties[".tag"].Value -eq "folder") { 0 } else { 1 } }, name
    if ($entries) {
        $global:dropboxEntries = @{}
        foreach ($entry in $entries) {
            if ($entry.name) {
                $lvi = New-Object System.Windows.Forms.ListViewItem($entry.name)
                if ($entry.PSObject.Properties[".tag"].Value -eq "folder") {
                    $lvi.ImageKey = "folder"
                } else {
                    # Para archivos, obtenemos la extensión y usamos el ícono asociado
                    $ext = [System.IO.Path]::GetExtension($entry.name).ToLower()
                    if ([string]::IsNullOrEmpty($ext)) { $ext = "file" }
                    if (-not $imageList.Images.ContainsKey($ext)) {
                        try {
                            # Se utiliza un nombre ficticio para extraer el ícono asociado a la extensión
                            $dummyFile = "dummy" + $ext
                            $iconBmp = Get-FileIcon -filePath $dummyFile
                            $imageList.Images.Add($ext, $iconBmp)
                        } catch {
                            $ext = "file"
                        }
                    }
                    $lvi.ImageKey = $ext
                }
                $fileListView.Items.Add($lvi)
                $global:dropboxEntries[$entry.name] = $entry
            }
        }
    }
}

$fileListView.Add_DoubleClick({
    if ($fileListView.SelectedItems.Count -gt 0) {
        $selectedItem = $fileListView.SelectedItems[0].Text
        Debug-Print "Double-clicked item: '$selectedItem'"
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
            $entry = $global:dropboxEntries[$selectedItem]
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
})

$global:accessToken = Get-AccessToken
if ($global:accessToken) {
    Update-FileList
    $form.ShowDialog()
} else {
    Write-Host "No se pudo obtener el token. Verifica tus credenciales."
}
