Add-Type -AssemblyName System.Windows.Forms

# Función para obtener el token de acceso
def function Get-AccessToken {
    $AppKey = "8g8oqwp5x26h58o"
    $AppSecret = "z3690pwzqowtzjx"
    $RefreshToken = "g8E8wsPIxW8AAAAAAAAAAUVpLeEobmxg1sWlIdufgjninvxJp2x4-YLIC53n6gNe"

    $Body = @{ refresh_token = $RefreshToken; grant_type = "refresh_token"; client_id = $AppKey; client_secret = $AppSecret }
    $Headers = @{ "Content-Type" = "application/x-www-form-urlencoded" }
    $Url = "https://api.dropboxapi.com/oauth2/token"
    
    try {
        $Response = Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body
        return $Response.access_token
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al renovar el token: $($_.Exception.Message)")
        return $null
    }
}

# Configuración y variables globales
$token = Get-AccessToken
$headers = @{ "Authorization" = "Bearer $token"; "Content-Type" = "application/json" }
$global:downloadDirectory = ""
$global:currentPath = ""
$global:history = @()
$global:downloadedFiles = @()

# Crear Formulario
$form = New-Object System.Windows.Forms.Form
$form.Text = "Explorador de Dropbox"
$form.Size = New-Object System.Drawing.Size(500, 500)

# ListBox para carpetas
$listBoxFolders = New-Object System.Windows.Forms.ListBox
$listBoxFolders.Location = New-Object System.Drawing.Point(10, 10)
$listBoxFolders.Size = New-Object System.Drawing.Size(200, 200)
$form.Controls.Add($listBoxFolders)

# ListBox para archivos
$listBoxFiles = New-Object System.Windows.Forms.ListBox
$listBoxFiles.Location = New-Object System.Drawing.Point(220, 10)
$listBoxFiles.Size = New-Object System.Drawing.Size(200, 200)
$form.Controls.Add($listBoxFiles)

# ListBox para mostrar estado de descargas
$listBoxStatus = New-Object System.Windows.Forms.ListBox
$listBoxStatus.Location = New-Object System.Drawing.Point(10, 220)
$listBoxStatus.Size = New-Object System.Drawing.Size(410, 100)
$form.Controls.Add($listBoxStatus)

# Botón para seleccionar carpeta de descarga
$btnSelectFolder = New-Object System.Windows.Forms.Button
$btnSelectFolder.Text = "Seleccionar Carpeta"
$btnSelectFolder.Location = New-Object System.Drawing.Point(10, 330)
$btnSelectFolder.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq "OK") {
        $global:downloadDirectory = $folderDialog.SelectedPath
        $listBoxStatus.Items.Add("Carpeta de descarga: $global:downloadDirectory")
    }
})
$form.Controls.Add($btnSelectFolder)

# Botón para descargar archivos seleccionados
$btnDownloadFiles = New-Object System.Windows.Forms.Button
$btnDownloadFiles.Text = "Descargar Archivos"
$btnDownloadFiles.Location = New-Object System.Drawing.Point(150, 330)
$btnDownloadFiles.Add_Click({
    if ($global:downloadDirectory -eq "") {
        [System.Windows.Forms.MessageBox]::Show("Seleccione una carpeta de descarga primero.")
        return
    }
    
    $selectedFiles = $listBoxFiles.SelectedItems
    if ($selectedFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Seleccione al menos un archivo.")
        return
    }

    foreach ($file in $selectedFiles) {
        Download-File $file
    }

    Show-SortDialog
})
$form.Controls.Add($btnDownloadFiles)

# Función para actualizar listas
function Update-Lists {
    $listBoxFolders.Items.Clear()
    $listBoxFiles.Items.Clear()
    $listBoxStatus.Items.Clear()
    
    if ($global:currentPath -ne "") {
        $listBoxFolders.Items.Add("...")
    }
    
    $folders = Get-Folders
    foreach ($folder in $folders) { 
        $listBoxFolders.Items.Add($folder.name)
    }
    
    $files = Get-Files
    foreach ($file in $files) {
        $listBoxFiles.Items.Add($file.name)
    }
}

# Evento de doble clic para cambiar de carpeta
$listBoxFolders.Add_DoubleClick({
    if ($listBoxFolders.SelectedItem -eq "...") {
        $global:currentPath = $global:history[-1]
        $global:history = $global:history[0..($global:history.Count - 2)]
    } else {
        $global:history += $global:currentPath
        $global:currentPath += "/" + $listBoxFolders.SelectedItem
    }
    Update-Lists
})

# Actualizar listas antes de mostrar el formulario
Update-Lists
$form.ShowDialog()
