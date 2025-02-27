Add-Type -AssemblyName System.Windows.Forms

# Función para obtener el token de acceso
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
        $Response = Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body
        return $Response.access_token
    } catch {
        Write-Host "Error al renovar el token: $($_.Exception.Message)"
        return $null
    }
}

# Obtener el token de acceso
$global:accessToken = Get-AccessToken
if (-not $global:accessToken) {
    [System.Windows.Forms.MessageBox]::Show("Error al obtener el token de acceso.")
    exit
}

# Crear formulario principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Explorador de Dropbox"
$form.Size = New-Object System.Drawing.Size(600, 450)

# Crear ListBox para carpetas
$global:listBoxFolders = New-Object System.Windows.Forms.ListBox
$global:listBoxFolders.Location = New-Object System.Drawing.Point(10, 10)
$global:listBoxFolders.Size = New-Object System.Drawing.Size(260, 150)
$global:listBoxFolders.SelectionMode = "MultiExtended"
$global:listBoxFolders.Add_DoubleClick({ Enter-Folder })
$form.Controls.Add($global:listBoxFolders)

# Crear ListBox para archivos
$global:listBoxFiles = New-Object System.Windows.Forms.ListBox
$global:listBoxFiles.Location = New-Object System.Drawing.Point(280, 10)
$global:listBoxFiles.Size = New-Object System.Drawing.Size(260, 150)
$global:listBoxFiles.SelectionMode = "MultiExtended"
$form.Controls.Add($global:listBoxFiles)

# Crear ListBox para estado de descargas
$global:listBoxStatus = New-Object System.Windows.Forms.ListBox
$global:listBoxStatus.Location = New-Object System.Drawing.Point(10, 170)
$global:listBoxStatus.Size = New-Object System.Drawing.Size(530, 100)
$form.Controls.Add($global:listBoxStatus)

# Botón para seleccionar carpeta de descarga
$btnSelectFolder = New-Object System.Windows.Forms.Button
$btnSelectFolder.Text = "Seleccionar Carpeta"
$btnSelectFolder.Location = New-Object System.Drawing.Point(10, 280)
$btnSelectFolder.Size = New-Object System.Drawing.Size(150, 30)
$btnSelectFolder.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $global:downloadDirectory = $folderDialog.SelectedPath
    }
})
$form.Controls.Add($btnSelectFolder)

# Botón para descargar archivos seleccionados
$btnDownload = New-Object System.Windows.Forms.Button
$btnDownload.Text = "Descargar"
$btnDownload.Location = New-Object System.Drawing.Point(180, 280)
$btnDownload.Size = New-Object System.Drawing.Size(100, 30)
$btnDownload.Add_Click({ DownloadFiles })
$form.Controls.Add($btnDownload)

# Función para obtener carpetas y archivos
function Update-Lists {
    $global:listBoxFolders.Items.Clear()
    $global:listBoxFiles.Items.Clear()
    
    if ($global:currentPath -ne "") {
        $global:listBoxFolders.Items.Add("...")
    }
    
    $url = "https://api.dropboxapi.com/2/files/list_folder"
    $body = @{ path = $global:currentPath } | ConvertTo-Json -Compress
    $headers = @{ "Authorization" = "Bearer $global:accessToken"; "Content-Type" = "application/json" }
    
    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        foreach ($entry in $response.entries) {
            if ($entry.".tag" -eq "folder") {
                $global:listBoxFolders.Items.Add($entry.name)
            } elseif ($entry.".tag" -eq "file") {
                $global:listBoxFiles.Items.Add($entry.name)
            }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al cargar archivos y carpetas.")
    }
}

# Función para descargar archivos de Dropbox
function Download-FileFromDropbox($fileName) {
    $dropboxPath = if ($global:currentPath -eq "") { "/$fileName" } else { "$global:currentPath/$fileName" }
    $localPath = Join-Path $global:downloadDirectory $fileName

    $url = "https://content.dropboxapi.com/2/files/download"
    $headers = @{
        "Authorization"    = "Bearer $global:accessToken"
        "Dropbox-API-Arg"  = ("{`"path`": `"$dropboxPath`"}" | ConvertTo-Json -Compress).Replace('"', '\"')
    }

    try {
        Invoke-RestMethod -Uri $url -Method Get -Headers $headers -OutFile $localPath
        return $localPath
    } catch {
        Write-Host "Error al descargar $fileName $($_.Exception.Message)"
        return $null
    }
}

# Función para descargar archivos antes de ordenar
function DownloadFiles {
    if ([string]::IsNullOrEmpty($global:downloadDirectory)) {
        [System.Windows.Forms.MessageBox]::Show("Selecciona primero la carpeta de destino.")
        return
    }
    
    $selectedFiles = $global:listBoxFiles.SelectedItems
    if ($selectedFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No se ha seleccionado ningún archivo.")
        return
    }
    
    $global:downloadedFiles = @()
    
    foreach ($file in $selectedFiles) {
        $downloadedFile = Download-FileFromDropbox $file
        if ($downloadedFile) {
            $global:listBoxStatus.Items.Add("$file descargado correctamente.")
            $global:downloadedFiles += $downloadedFile
        } else {
            $global:listBoxStatus.Items.Add("Error al descargar $file.")
        }
    }
    
    Open-SortWindow
}

# Ventana de ordenamiento de archivos
function Open-SortWindow {
    $sortForm = New-Object System.Windows.Forms.Form
    $sortForm.Text = "Ordenar Archivos"
    $sortForm.Size = New-Object System.Drawing.Size(400, 500)
    
    $listBoxSort = New-Object System.Windows.Forms.ListBox
    $listBoxSort.Location = New-Object System.Drawing.Point(10, 10)
    $listBoxSort.Size = New-Object System.Drawing.Size(360, 250)
    foreach ($file in $global:downloadedFiles) {
        $listBoxSort.Items.Add($file)
    }
    $sortForm.Controls.Add($listBoxSort)

    $sortForm.ShowDialog()
}

Update-Lists
$form.ShowDialog()

