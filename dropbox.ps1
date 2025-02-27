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
        [System.Windows.Forms.MessageBox]::Show("Error al renovar el token: $($_.ErrorDetails)")
        return $null
    }
}

# Obtener el token de acceso
$global:accessToken = Get-AccessToken
if (-not $global:accessToken) {
    [System.Windows.Forms.MessageBox]::Show("Error al obtener el token de acceso.")
    exit
}

# Variables globales
$global:currentPath = ""
$global:selectedFiles = @()

# Función para listar archivos y carpetas
function Load-DropboxContent {
    param ($path)
    $global:listBoxFolders.Items.Clear()
    $global:currentPath = $path

    if ($path -ne "") {
        $global:listBoxFolders.Items.Add("...")
    }

    $headers = @{ "Authorization" = "Bearer $global:accessToken"; "Content-Type" = "application/json" }
    $body = @{ path = $path } | ConvertTo-Json -Depth 10 -Compress
    $url = "https://api.dropboxapi.com/2/files/list_folder"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        if ($response.entries.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No se encontraron archivos en esta carpeta.")
        } else {
            foreach ($item in $response.entries) {
                $global:listBoxFolders.Items.Add($item.name)
            }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al cargar contenido de Dropbox: $($_.ErrorDetails)")
    }
}

# Función para entrar a carpetas o seleccionar archivos
function Enter-Item {
    $selectedItem = $global:listBoxFolders.SelectedItem
    if ($null -eq $selectedItem) { return }

    if ($selectedItem -eq "...") {
        $global:currentPath = ($global:currentPath -split "/")[0..($global:currentPath.Split("/").Count - 2)] -join "/"
        Load-DropboxContent -path $global:currentPath
        return
    }

    $headers = @{ "Authorization" = "Bearer $global:accessToken"; "Content-Type" = "application/json" }
    $body = @{ path = "$global:currentPath/$selectedItem" } | ConvertTo-Json -Compress
    $url = "https://api.dropboxapi.com/2/files/get_metadata"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        if ($response.".tag" -eq "folder") {
            Load-DropboxContent -path "$global:currentPath/$selectedItem"
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al acceder al ítem: $($_.ErrorDetails)")
    }
}

# Función para agregar archivos a la lista de descargas
function Move-ToDownloadList {
    $selectedItem = $global:listBoxFolders.SelectedItem
    if ($null -eq $selectedItem -or $selectedItem -eq "...") { return }
    
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
}

# Crear formulario principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Explorador de Dropbox"
$form.Size = New-Object System.Drawing.Size(800, 450)

# ListBox para mostrar carpetas y archivos
$global:listBoxFolders = New-Object System.Windows.Forms.ListBox
$global:listBoxFolders.Location = New-Object System.Drawing.Point(10, 10)
$global:listBoxFolders.Size = New-Object System.Drawing.Size(350, 350)
$global:listBoxFolders.Add_DoubleClick({ Enter-Item })
$form.Controls.Add($global:listBoxFolders)

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
