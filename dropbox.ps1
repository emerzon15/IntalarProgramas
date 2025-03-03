
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
        [System.Windows.Forms.MessageBox]::Show("Error al renovar el token: $($_.ErrorDetails)")
        return $null
    }
}
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
# Variables globales
$global:currentPath = ""
$global:selectedFiles = @()

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
# Función para listar archivos y carpetas
function Load-DropboxContent {
    param ($path)
    $global:listBoxFolders.Items.Clear()
    $global:listBoxFiles.Items.Clear()
    
    if ($global:currentPath -ne "") {
    $global:currentPath = $path

    if ($path -ne "") {
        $global:listBoxFolders.Items.Add("...")
    }
    
    $url = "https://api.dropboxapi.com/2/files/list_folder"
    $body = @{ path = $global:currentPath } | ConvertTo-Json -Compress

    $headers = @{ "Authorization" = "Bearer $global:accessToken"; "Content-Type" = "application/json" }
    
    $body = @{ path = $path } | ConvertTo-Json -Depth 10 -Compress
    $url = "https://api.dropboxapi.com/2/files/list_folder"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        foreach ($entry in $response.entries) {
            if ($entry.".tag" -eq "folder") {
                $global:listBoxFolders.Items.Add($entry.name)
            } elseif ($entry.".tag" -eq "file") {
                $global:listBoxFiles.Items.Add($entry.name)
        if ($response.entries.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No se encontraron archivos en esta carpeta.")
        } else {
            foreach ($item in $response.entries) {
                $global:listBoxFolders.Items.Add($item.name)
            }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al cargar archivos y carpetas.")
        [System.Windows.Forms.MessageBox]::Show("Error al cargar contenido de Dropbox: $($_.ErrorDetails)")
    }
}

# Función para descargar archivos de Dropbox
function Download-FileFromDropbox($fileName) {
    $dropboxPath = if ($global:currentPath -eq "") { "/$fileName" } else { "$global:currentPath/$fileName" }
    $localPath = Join-Path $global:downloadDirectory $fileName
# Función para entrar a carpetas o seleccionar archivos
function Enter-Item {
    $selectedItem = $global:listBoxFolders.SelectedItem
    if ($null -eq $selectedItem) { return }

    $url = "https://content.dropboxapi.com/2/files/download"
    $headers = @{
        "Authorization"    = "Bearer $global:accessToken"
        "Dropbox-API-Arg"  = ("{`"path`": `"$dropboxPath`"}" | ConvertTo-Json -Compress).Replace('"', '\"')
    if ($selectedItem -eq "...") {
        $global:currentPath = ($global:currentPath -split "/")[0..($global:currentPath.Split("/").Count - 2)] -join "/"
        Load-DropboxContent -path $global:currentPath
        return
    }

    $headers = @{ "Authorization" = "Bearer $global:accessToken"; "Content-Type" = "application/json" }
    $body = @{ path = "$global:currentPath/$selectedItem" } | ConvertTo-Json -Compress
    $url = "https://api.dropboxapi.com/2/files/get_metadata"

    try {
        Invoke-RestMethod -Uri $url -Method Get -Headers $headers -OutFile $localPath
        return $localPath
        $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        if ($response.".tag" -eq "folder") {
            Load-DropboxContent -path "$global:currentPath/$selectedItem"
        }
    } catch {
        Write-Host "Error al descargar $fileName $($_.Exception.Message)"
        return $null
        [System.Windows.Forms.MessageBox]::Show("Error al acceder al ítem: $($_.ErrorDetails)")
    }
}

# Función para descargar archivos antes de ordenar
function DownloadFiles {
    if ([string]::IsNullOrEmpty($global:downloadDirectory)) {
        [System.Windows.Forms.MessageBox]::Show("Selecciona primero la carpeta de destino.")
        return
    }
# Función para agregar archivos a la lista de descargas
function Move-ToDownloadList {
    $selectedItem = $global:listBoxFolders.SelectedItem
    if ($null -eq $selectedItem -or $selectedItem -eq "...") { return }

    $selectedFiles = $global:listBoxFiles.SelectedItems
    if ($selectedFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No se ha seleccionado ningún archivo.")
    $fullPath = "$global:currentPath/$selectedItem"

    if (-not $global:selectedFiles.Contains($fullPath)) {
        $global:selectedFiles += $fullPath
        $global:listBoxDownloads.Items.Add($selectedItem)
    }
}

# Función para eliminar archivos de la lista de descargas
function Remove-FromDownloadList {
    $selectedItem = $global:listBoxDownloads.SelectedItem
    if ($null -eq $selectedItem) { return }

    $global:selectedFiles = $global:selectedFiles | Where-Object { $_ -ne "$global:currentPath/$selectedItem" }
    $global:listBoxDownloads.Items.Remove($selectedItem)
}

# Función para descargar archivos desde Dropbox
function Download-Files {
    if ($global:selectedFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No hay archivos seleccionados para descargar.")
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

    $downloadDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($downloadDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $localPath = $downloadDialog.SelectedPath

    foreach ($file in $global:selectedFiles) {
        $filePath = if ($file.StartsWith("/")) { $file } else { "/$file" }

        $headers = @{ 
            "Authorization" = "Bearer $global:accessToken"
            "Dropbox-API-Arg" = ('{"path": "' + $filePath + '"}')
        }
        $url = "https://content.dropboxapi.com/2/files/download"
        $localFilePath = Join-Path -Path $localPath -ChildPath (Split-Path -Leaf $filePath)

        try {
            Invoke-WebRequest -Uri $url -Method Post -Headers $headers -OutFile $localFilePath
            [System.Windows.Forms.MessageBox]::Show("Archivo descargado: $localFilePath")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al descargar $filePath`n$($_.Exception.Message)")
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
# Crear formulario principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Explorador de Dropbox"
$form.Size = New-Object System.Drawing.Size(800, 450)

    $sortForm.ShowDialog()
}
# ListBox para mostrar carpetas y archivos
$global:listBoxFolders = New-Object System.Windows.Forms.ListBox
$global:listBoxFolders.Location = New-Object System.Drawing.Point(10, 10)
$global:listBoxFolders.Size = New-Object System.Drawing.Size(350, 350)
$global:listBoxFolders.Add_DoubleClick({ Enter-Item })
$form.Controls.Add($global:listBoxFolders)

Update-Lists
$form.ShowDialog()
# Botón "Mover a esa ventana"
$btnMove = New-Object System.Windows.Forms.Button
$btnMove.Text = "Mover a esa ventana"
$btnMove.Location = New-Object System.Drawing.Point(370, 10)
$btnMove.Add_Click({ Move-ToDownloadList })
$form.Controls.Add($btnMove)

# ListBox para archivos seleccionados
$global:listBoxDownloads = New-Object System.Windows.Forms.ListBox
$global:listBoxDownloads.Location = New-Object System.Drawing.Point(500, 10)
$global:listBoxDownloads.Size = New-Object System.Drawing.Size(250, 250)
$form.Controls.Add($global:listBoxDownloads)

# Botón "Eliminar de lista"
$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Eliminar"
$btnRemove.Location = New-Object System.Drawing.Point(500, 270)
$btnRemove.Add_Click({ Remove-FromDownloadList })
$form.Controls.Add($btnRemove)

# Botón "Descargar"
$btnDownload = New-Object System.Windows.Forms.Button
$btnDownload.Text = "Descargar"
$btnDownload.Location = New-Object System.Drawing.Point(500, 310)
$btnDownload.Add_Click({ Download-Files })
$form.Controls.Add($btnDownload)

# Cargar la raíz de Dropbox
Load-DropboxContent -path ""

# Mostrar el formulario
$form.ShowDialog()
