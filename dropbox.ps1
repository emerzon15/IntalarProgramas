# ===============================================================
# ====================== INICIO DEL CÓDIGO E.C.N=================
# ===============================================================

# Cargar ensamblados necesarios
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

# --- Función para obtener el ícono asociado a una extensión ---
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

# --- Función para determinar si el usuario actual es administrador ---
function Is-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# --- Depuración ---
$global:DebugMode = $true
function Debug-Print($message) {
    if ($global:DebugMode) { Write-Host "[DEBUG] $message" }
}

# ========================================================
# == Funciones de Dropbox (API y Descarga/Upload) ==
# ========================================================

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

# -------------------------------
#   NUEVAS FUNCIONES CON PROGRESO DE SUBIDA
# -------------------------------

# Subida de un archivo con "falso" progreso: 0% → 100%
function Upload-DropboxFileNoExecute($LocalFilePath, $DropboxDestinationFolder, $panel) {
    $FileName = [System.IO.Path]::GetFileName($LocalFilePath)

    # Barra y label del panel (ya creados antes de llamar esta función)
    $progressBar = $panel.Controls | Where-Object { $_ -is [System.Windows.Forms.ProgressBar] }
    $statusLabel = $panel.Controls | Where-Object { $_ -is [System.Windows.Forms.Label] -and $_.Name -eq "StatusLabel" }

    # Iniciamos en 0%
    $progressBar.Value = 0
    $statusLabel.Text  = "0%"

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
        # Subida sin progreso parcial real (bloque único)
        Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -InFile $LocalFilePath -ContentType "application/octet-stream" | Out-Null

        # Si todo va bien, lo ponemos en 100% / Completado
        $progressBar.Value = 100
        $statusLabel.Text  = "100% / Completado"
        return $true
    } catch {
        Debug-Print "Error al subir archivo: $($_.Exception.Message)"
        $statusLabel.Text = "Error al subir"
        return $false
    }
}

# Subida recursiva de carpeta con "falso" progreso (por archivo)
function Upload-DropboxFolderNoExecute($LocalFolderPath, $DropboxDestinationFolder, $parentPanel) {
    $resultPaths = @()
    $FolderName = [System.IO.Path]::GetFileName($LocalFolderPath)
    if (-not $FolderName) { $FolderName = "FolderLocal" }

    # Creamos un panel con barra de progreso para la carpeta
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(([int]$parentPanel.Width - 25), 40)
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $folderLabel = New-Object System.Windows.Forms.Label
    $folderLabel.Text = "Subiendo carpeta: $FolderName"
    $folderLabel.AutoSize = $false
    $folderLabel.AutoEllipsis = $true
    $folderLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $folderLabel.Size = New-Object System.Drawing.Size(400,20)
    $folderLabel.Location = New-Object System.Drawing.Point(5,5)
    $panel.Controls.Add($folderLabel) | Out-Null

    $folderProgressBar = New-Object System.Windows.Forms.ProgressBar
    $folderProgressBar.Location = New-Object System.Drawing.Point(410,5)
    $folderProgressBar.Size = New-Object System.Drawing.Size(250,20)
    $folderProgressBar.Style = 'Continuous'
    $panel.Controls.Add($folderProgressBar) | Out-Null
    $panel.Refresh()
    $null = $folderProgressBar.Handle
    [ThemeHelper]::SetWindowTheme($folderProgressBar.Handle, "", "")
    [Win32]::SendMessage($folderProgressBar.Handle, $PBM_SETBARCOLOR, [IntPtr]::Zero, [System.Drawing.Color]::Red.ToArgb())

    $folderStatusLabel = New-Object System.Windows.Forms.Label
    $folderStatusLabel.Name = "StatusLabel"  # para identificarla
    $folderStatusLabel.Text = "0%"
    $folderStatusLabel.Size = New-Object System.Drawing.Size(120,20)
    $folderStatusLabel.Location = New-Object System.Drawing.Point(665,5)
    $panel.Controls.Add($folderStatusLabel) | Out-Null

    $parentPanel.Controls.Add($panel) | Out-Null
    $parentPanel.ScrollControlIntoView($panel)
    $parentPanel.Refresh()

    # Primero creamos la carpeta en Dropbox
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

    # Recorrer todos los archivos y subcarpetas
    $allFiles = Get-ChildItem -Path $LocalFolderPath -Recurse -File
    $totalFiles = $allFiles.Count
    $filesUploaded = 0

    foreach ($file in $allFiles) {
        # Calcular ruta de Dropbox para cada subarchivo
        $relativePath = $file.FullName.Substring($LocalFolderPath.Length).TrimStart("\")
        $destinationPath = Join-Path $DropboxFolderPath $relativePath

        # Crear un panel de "archivo" para cada uno, o podemos subirlos "silenciosamente".
        # Aquí haremos "silencioso" para no crear decenas de paneles:
        # Si quieres un panel por cada archivo, puedes hacerlo igual que en Upload-DropboxFileNoExecute.
        # Por simplicidad, subimos en bloque y solo actualizamos la barra "carpeta".
        try {
            $UrlFile = "https://content.dropboxapi.com/2/files/upload"
            $HeadersFile = @{
                "Authorization" = "Bearer $global:accessToken"
                "Dropbox-API-Arg" = (ConvertTo-Json -Compress @{
                    path = $destinationPath
                    mode = "add"
                    autorename = $true
                    mute = $false
                    strict_conflict = $false
                })
                "Content-Type" = "application/octet-stream"
            }
            Invoke-RestMethod -Uri $UrlFile -Method Post -Headers $HeadersFile -InFile $file.FullName -ContentType "application/octet-stream" | Out-Null
            $filesUploaded++

            # Actualizamos la barra para la carpeta
            $percentage = [math]::Round(($filesUploaded / $totalFiles * 100))
            $folderProgressBar.Value = $percentage
            $folderStatusLabel.Text  = "$percentage%"

        } catch {
            Debug-Print "Error al subir archivo: $($_.Exception.Message)"
        }
    }

    if ($filesUploaded -gt 0) {
        $folderProgressBar.Value = 100
        $folderStatusLabel.Text = "100% / Completado"
    } else {
        $folderStatusLabel.Text += " / Sin archivos"
    }

    return $resultPaths
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

# -------------------------------
# Funciones de Descarga (Sin Execute)
# -------------------------------

function Download-DropboxFileWithProgressNoExecute($FilePath, $LocalPath, $parentPanel) {
    $FileName = [System.IO.Path]::GetFileName($FilePath)
    if (-not $FileName) {
        Debug-Print "No se pudo determinar nombre para '$FilePath'"
        return $null
    }
    $OutFilePath = Join-Path -Path $LocalPath -ChildPath $FileName

    # Panel para cada archivo en descarga
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(([int]$parentPanel.Width - 25), 40)
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    
    # Label para el nombre del archivo
    $fileNameLabel = New-Object System.Windows.Forms.Label
    $fileNameLabel.Text = $FileName
    $fileNameLabel.AutoSize = $false
    $fileNameLabel.AutoEllipsis = $true
    $fileNameLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $fileNameLabel.Size = New-Object System.Drawing.Size(400,20)
    $fileNameLabel.Location = New-Object System.Drawing.Point(5,5)
    $panel.Controls.Add($fileNameLabel) | Out-Null

    # Barra de progreso
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(410,5)
    $progressBar.Size = New-Object System.Drawing.Size(250,20)
    $progressBar.Style = 'Continuous'
    $panel.Controls.Add($progressBar) | Out-Null
    $panel.Refresh()
    $null = $progressBar.Handle
    [ThemeHelper]::SetWindowTheme($progressBar.Handle, "", "")
    [Win32]::SendMessage($progressBar.Handle, $PBM_SETBARCOLOR, [IntPtr]::Zero, [System.Drawing.Color]::Red.ToArgb())

    # Label de estado (porcentaje, completado, etc.)
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "0%"
    $statusLabel.Size = New-Object System.Drawing.Size(100,20)
    $statusLabel.Location = New-Object System.Drawing.Point(665,5)
    $panel.Controls.Add($statusLabel) | Out-Null

    $parentPanel.Controls.Add($panel) | Out-Null
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
            $progressBar.Value = 100
            $statusLabel.Text = "100% / Completado"
        }
        if ($null -ne $done) { $done.Set() | Out-Null }
    })

    try {
        $uri = [System.Uri] "https://content.dropboxapi.com/2/files/download"
        $webClient.DownloadFileAsync($uri, $OutFilePath)
    } catch {
        $statusLabel.Text = "Error al iniciar descarga"
        Debug-Print "Error al iniciar descarga: $($_.Exception.Message)"
        if ($null -ne $done) { $done.Set() | Out-Null }
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
        if ($null -ne $done) { $done.Set() | Out-Null }
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

# ======================
# Función de Descarga de Carpeta (modificada para descarga recursiva)
# ======================
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
    
    # Panel de progreso para la carpeta
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(([int]$parentPanel.Width - 25), 40)
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $folderLabel = New-Object System.Windows.Forms.Label
    $folderLabel.Text = "Descargando carpeta: $folderName"
    $folderLabel.AutoSize = $false
    $folderLabel.AutoEllipsis = $true
    $folderLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $folderLabel.Size = New-Object System.Drawing.Size(400,20)
    $folderLabel.Location = New-Object System.Drawing.Point(5,5)
    $panel.Controls.Add($folderLabel) | Out-Null

    $folderProgressBar = New-Object System.Windows.Forms.ProgressBar
    $folderProgressBar.Location = New-Object System.Drawing.Point(410,5)
    $folderProgressBar.Size = New-Object System.Drawing.Size(250,20)
    $folderProgressBar.Style = 'Continuous'
    $panel.Controls.Add($folderProgressBar) | Out-Null
    $panel.Refresh()
    $null = $folderProgressBar.Handle
    [ThemeHelper]::SetWindowTheme($folderProgressBar.Handle, "", "")
    [Win32]::SendMessage($folderProgressBar.Handle, $PBM_SETBARCOLOR, [IntPtr]::Zero, [System.Drawing.Color]::Red.ToArgb())

    $folderStatusLabel = New-Object System.Windows.Forms.Label
    $folderStatusLabel.Text = "0%"
    $folderStatusLabel.Size = New-Object System.Drawing.Size(100,20)
    $folderStatusLabel.Location = New-Object System.Drawing.Point(665,5)
    $panel.Controls.Add($folderStatusLabel) | Out-Null

    $parentPanel.Controls.Add($panel) | Out-Null
    $parentPanel.ScrollControlIntoView($panel)
    $parentPanel.Refresh()

    $filesDownloaded = 0

    function DownloadFilesRecursively {
        param(
            [string]$currentDropboxPath,
            [string]$currentLocalPath
        )
        $entries = Get-DropboxFiles -Path $currentDropboxPath
        $totalEntries = $entries.Count
        foreach ($entry in $entries) {
            if ($entry.PSObject.Properties[".tag"].Value -eq "folder") {
                # Crear subcarpeta local y llamar recursivamente
                $subLocalFolder = Join-Path $currentLocalPath $entry.name
                if (!(Test-Path $subLocalFolder)) {
                    New-Item -ItemType Directory -Path $subLocalFolder | Out-Null
                }
                DownloadFilesRecursively -currentDropboxPath $entry.path_lower -currentLocalPath $subLocalFolder
            } else {
                $localPath = Download-DropboxFileSilent $entry.path_lower $currentLocalPath
                if ($localPath) {
                    $resultPaths += $localPath
                    $filesDownloaded++
                    # Actualizar progreso para esta carpeta
                    $percentage = [math]::Round(($filesDownloaded / $totalEntries * 100))
                    $folderProgressBar.Value = $percentage
                    $folderStatusLabel.Text  = "$percentage%"
                }
            }
        }
    }

    DownloadFilesRecursively -currentDropboxPath $FolderPath -currentLocalPath $localFolderPath
    if ($filesDownloaded -gt 0) {
        $folderProgressBar.Value = 100
        $folderStatusLabel.Text = "100% / Completado"
    } else {
        $folderStatusLabel.Text += " / Sin archivos"
    }
    return $resultPaths
}

# ==========================================================
# == INTERFAZ GRÁFICA (CÓDIGO COMPLETO UNIDO Y MODIFICADO) ==
# ==========================================================

# Crear formulario principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "UTP"
$form.Size = New-Object System.Drawing.Size(900,600)
$form.StartPosition = "CenterScreen"

# Crear ImageList e íconos
$imageList = New-Object System.Windows.Forms.ImageList
$folderIcon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\explorer.exe").ToBitmap()
$imageList.Images.Add("folder", $folderIcon)
# Ícono por defecto para archivos
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
$fileListView.MultiSelect = $true    # Habilitar selección múltiple (Ctrl, Shift)
$null = $form.Controls.Add($fileListView)

# ListBox (derecha) para ítems seleccionados para descarga
$selectedFilesBox = New-Object System.Windows.Forms.ListBox
$selectedFilesBox.Location = New-Object System.Drawing.Point(460,10)
$selectedFilesBox.Size = New-Object System.Drawing.Size(350,400)
$null = $form.Controls.Add($selectedFilesBox)

# Función para aplicar estilo a botones
function Set-ButtonStyle($button) {
    $button.BackColor  = [System.Drawing.Color]::Red
    $button.ForeColor  = [System.Drawing.Color]::White
    $button.FlatStyle  = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $button.FlatAppearance.BorderSize  = 1
    $button.TextAlign  = [System.Drawing.ContentAlignment]::MiddleCenter
}

# Botón para pasar ítems de la izquierda a la derecha (a lista de descarga)
$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = "→"
$addButton.Size = New-Object System.Drawing.Size(30,30)
$addButton.Location = New-Object System.Drawing.Point(420,150)
Set-ButtonStyle $addButton
$addButton.Add_Click({
    foreach ($selectedItem in $fileListView.SelectedItems) {
        $itemText = $selectedItem.Text
        if ($itemText -and $itemText -ne "..") {
            $entry = $global:dropboxEntries[$itemText]
            $filePath = if ($entry) { $entry.path_lower } else { "$global:currentPath/$itemText" -replace "//","/" }
            if (-not $selectedFilesBox.Items.Contains($filePath)) {
                $selectedFilesBox.Items.Add($filePath) | Out-Null
            }
        }
    }
})
$null = $form.Controls.Add($addButton)

$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = "←"
$removeButton.Size = New-Object System.Drawing.Size(30,30)
$removeButton.Location = New-Object System.Drawing.Point(420,190)
Set-ButtonStyle $removeButton
$removeButton.Add_Click({
    $selectedFilesBox.Items.Remove($selectedFilesBox.SelectedItem) | Out-Null
})
$null = $form.Controls.Add($removeButton)

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
        $selectedFilesBox.Items.Insert($index - 1, $item) | Out-Null
        $selectedFilesBox.SelectedIndex = $index - 1
    }
})
$null = $form.Controls.Add($moveUpButton)

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
        $selectedFilesBox.Items.Insert($index + 1, $item) | Out-Null
        $selectedFilesBox.SelectedIndex = $index + 1
    }
})
$null = $form.Controls.Add($moveDownButton)

# Panel para mostrar el progreso de descarga y subida
$downloadStatusPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$downloadStatusPanel.Location = New-Object System.Drawing.Point(10,500)
$downloadStatusPanel.Size = New-Object System.Drawing.Size(870,60)
$downloadStatusPanel.AutoScroll = $true
$null = $form.Controls.Add($downloadStatusPanel)
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
    $uploadForm.Controls.Add($lbl) | Out-Null

    $btnArchivo = New-Object System.Windows.Forms.Button
    $btnArchivo.Text = "Archivo"
    $btnArchivo.Size = New-Object System.Drawing.Size(70,30)
    $btnArchivo.Location = New-Object System.Drawing.Point(10,50)
    $btnArchivo.Add_Click({
        $uploadForm.Tag = "file"
        $uploadForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $uploadForm.Close()
    })
    $uploadForm.Controls.Add($btnArchivo) | Out-Null

    $btnCarpeta = New-Object System.Windows.Forms.Button
    $btnCarpeta.Text = "Carpeta"
    $btnCarpeta.Size = New-Object System.Drawing.Size(70,30)
    $btnCarpeta.Location = New-Object System.Drawing.Point(90,50)
    $btnCarpeta.Add_Click({
        $uploadForm.Tag = "folder"
        $uploadForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $uploadForm.Close()
    })
    $uploadForm.Controls.Add($btnCarpeta) | Out-Null

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Size = New-Object System.Drawing.Size(70,30)
    $btnCancelar.Location = New-Object System.Drawing.Point(170,50)
    $btnCancelar.Add_Click({
        $uploadForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $uploadForm.Close()
    })
    $uploadForm.Controls.Add($btnCancelar) | Out-Null

    $dialogResult = $uploadForm.ShowDialog($form)
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $option = $uploadForm.Tag
        if ($option -eq "file") {
            $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $fileDialog.Multiselect = $true
            if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                foreach ($file in $fileDialog.FileNames) {
                    # Crear panel de estado para la subida del archivo
                    $panel = New-Object System.Windows.Forms.Panel
                    $panel.Size = New-Object System.Drawing.Size(([int]$global:downloadStatusPanel.Width - 25), 40)
                    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

                    $fileNameLabel = New-Object System.Windows.Forms.Label
                    $fileNameLabel.Text = [System.IO.Path]::GetFileName($file)
                    $fileNameLabel.AutoSize = $false
                    $fileNameLabel.AutoEllipsis = $true
                    $fileNameLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $fileNameLabel.Size = New-Object System.Drawing.Size(200,20)
                    $fileNameLabel.Location = New-Object System.Drawing.Point(5,5)
                    $panel.Controls.Add($fileNameLabel) | Out-Null

                    $progressBar = New-Object System.Windows.Forms.ProgressBar
                    $progressBar.Location = New-Object System.Drawing.Point(210,5)
                    $progressBar.Size = New-Object System.Drawing.Size(250,20)
                    $progressBar.Style = 'Continuous'
                    $panel.Controls.Add($progressBar) | Out-Null
                    $panel.Refresh()
                    $null = $progressBar.Handle
                    [ThemeHelper]::SetWindowTheme($progressBar.Handle, "", "")
                    [Win32]::SendMessage($progressBar.Handle, $PBM_SETBARCOLOR, [IntPtr]::Zero, [System.Drawing.Color]::Red.ToArgb())

                    $statusLabel = New-Object System.Windows.Forms.Label
                    $statusLabel.Name = "StatusLabel"
                    $statusLabel.Text = "Subiendo..."
                    $statusLabel.Size = New-Object System.Drawing.Size(150,20)
                    $statusLabel.Location = New-Object System.Drawing.Point(465,5)
                    $panel.Controls.Add($statusLabel) | Out-Null

                    $global:downloadStatusPanel.Controls.Add($panel) | Out-Null
                    $global:downloadStatusPanel.ScrollControlIntoView($panel)
                    
                    # Usamos la nueva función con progreso "falso"
                    $result = Upload-DropboxFileNoExecute $file $global:currentPath $panel
                    if (-not $result) {
                        # Si falla, el label ya dirá "Error al subir"
                    }
                }
                Update-FileList
            }
        } elseif ($option -eq "folder") {
            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                # Subir carpeta con progreso
                Upload-DropboxFolderNoExecute $folderDialog.SelectedPath $global:currentPath $global:downloadStatusPanel
                Update-FileList
            }
        }
    }
})
$null = $form.Controls.Add($uploadButton)

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
$null = $form.Controls.Add($createFolderButton)

# Botón "Descargar": descarga todos los ítems y luego los ejecuta en orden.
$downloadButton = New-Object System.Windows.Forms.Button
$downloadButton.Text = "Descargar"
$downloadButton.Size = New-Object System.Drawing.Size(100,40)
$downloadButton.Location = New-Object System.Drawing.Point(230,420)
Set-ButtonStyle $downloadButton
$downloadButton.Add_Click({
    # Se fija la carpeta de Descargas por defecto en el usuario actual
    $chosenPath = Join-Path -Path $env:USERPROFILE -ChildPath "Downloads"
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
    # Filtrar únicamente rutas existentes para evitar errores
    $allDownloadedPaths = $allDownloadedPaths | Where-Object { Test-Path $_ }
    
    # Definir credenciales del usuario Administrador (se usarán solo si no somos admin)
    $adminUser = ".\Administrador"
    $adminPass = "UTPL4b$AQP265" | ConvertTo-SecureString -AsPlainText -Force
    $adminCred = New-Object System.Management.Automation.PSCredential($adminUser, $adminPass)
    
    foreach ($localFile in $allDownloadedPaths) {
        Debug-Print "Desbloqueando y ejecutando: $localFile"
        try {
            Unblock-File -Path $localFile
            if (Is-Admin) {
                Start-Process -FilePath $localFile -WindowStyle Hidden
            }
            else {
                Start-Process -FilePath $localFile -WindowStyle Hidden -Credential $adminCred
            }
        } catch {
            Debug-Print "No se pudo ejecutar $localFile $($_.Exception.Message)"
        }
    }
})
$null = $form.Controls.Add($downloadButton)

# Menú contextual para el ListView (Copiar, Cortar, Pegar, Eliminar)
$global:clipboardItem = @()
$global:clipboardOperation = $null
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuCopy   = New-Object System.Windows.Forms.ToolStripMenuItem("Copiar")
$menuCut    = New-Object System.Windows.Forms.ToolStripMenuItem("Cortar")
$menuPaste  = New-Object System.Windows.Forms.ToolStripMenuItem("Pegar")
$menuDelete = New-Object System.Windows.Forms.ToolStripMenuItem("Eliminar")
$contextMenu.Items.Add($menuCopy) | Out-Null
$contextMenu.Items.Add($menuCut) | Out-Null
$contextMenu.Items.Add($menuPaste) | Out-Null
$contextMenu.Items.Add($menuDelete) | Out-Null
$fileListView.ContextMenuStrip = $contextMenu

$menuCopy.Add_Click({
    if ($fileListView.SelectedItems.Count -gt 0) {
        $global:clipboardItem = @()
        foreach ($item in $fileListView.SelectedItems) {
            if ($item.Text -and $item.Text -ne "..") {
                $entry = $global:dropboxEntries[$item.Text]
                $global:clipboardItem += if ($entry) { $entry.path_lower } else { "$global:currentPath/$item.Text" -replace "//","/" }
            }
        }
        $global:clipboardOperation = "copy"
    }
})
$menuCut.Add_Click({
    if ($fileListView.SelectedItems.Count -gt 0) {
        $global:clipboardItem = @()
        foreach ($item in $fileListView.SelectedItems) {
            if ($item.Text -and $item.Text -ne "..") {
                $entry = $global:dropboxEntries[$item.Text]
                $global:clipboardItem += if ($entry) { $entry.path_lower } else { "$global:currentPath/$item.Text" -replace "//","/" }
            }
        }
        $global:clipboardOperation = "cut"
    }
})
$menuPaste.Add_Click({
    if ($global:clipboardItem -and $global:clipboardItem.Count -gt 0) {
        foreach ($item in $global:clipboardItem) {
            $destFile = [System.IO.Path]::GetFileName($item)
            $destination = if ([string]::IsNullOrEmpty($global:currentPath) -or $global:currentPath -eq "/") {
                "/$destFile"
            } else {
                "$global:currentPath/$destFile"
            }
            if ($global:clipboardOperation -eq "copy") {
                Copy-DropboxItem $item $destination | Out-Null
            } elseif ($global:clipboardOperation -eq "cut") {
                Move-DropboxItem $item $destination | Out-Null
            }
        }
        Update-FileList
        $global:clipboardItem = @()
        $global:clipboardOperation = $null
    }
})
$menuDelete.Add_Click({
    foreach ($item in $fileListView.SelectedItems) {
        if ($item.Text -and $item.Text -ne "..") {
            $entry = $global:dropboxEntries[$item.Text]
            $pathToDelete = if ($entry) { $entry.path_lower } else { "$global:currentPath/$item.Text" -replace "//","/" }
            Delete-DropboxItem $pathToDelete | Out-Null
        }
    }
    Update-FileList
})

$global:currentPath = ""
$global:dropboxEntries = @{}
function Update-FileList {
    if ($global:currentPath -eq "/") { $global:currentPath = "" }
    $fileListView.Items.Clear()
    if ($global:currentPath -ne "") {
        $upItem = New-Object System.Windows.Forms.ListViewItem("..")
        $upItem.ImageKey = "folder"
        $null = $fileListView.Items.Add($upItem)
    }
    # Se obtiene y ordena la lista de archivos y carpetas:
    # Primero las carpetas (valor 0) y luego los archivos (valor 1), ordenados alfabéticamente.
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
                            # Se usa un nombre ficticio para extraer el ícono asociado a la extensión
                            $dummyFile = "dummy" + $ext
                            $iconBmp = Get-FileIcon -filePath $dummyFile
                            $imageList.Images.Add($ext, $iconBmp)
                        } catch {
                            $ext = "file"
                        }
                    }
                    $lvi.ImageKey = $ext
                }
                $null = $fileListView.Items.Add($lvi)
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
    $form.ShowDialog() | Out-Null
} else {
    Write-Host "No se pudo obtener el token. Verifica tus credenciales."
}

# ===============================================================
# ======================= FIN DEL CÓDIGO ========================
# ===============================================================
# ==========================E.C.N================================
