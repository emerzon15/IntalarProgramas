dd-Type -AssemblyName System.Windows.Forms

# Habilitar mensajes de depuración
$global:DebugMode = $true
function Debug-Print($message) {
    if ($global:DebugMode) { Write-Host "[DEBUG] $message" }
}

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
        Debug-Print "Obteniendo token desde: $Url"
        $Response = Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body
        return $Response.access_token
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al renovar el token: $($_.Exception.Message)")
        return $null
    }
}

# Función para obtener la lista de archivos y carpetas de Dropbox
function Get-DropboxFiles($Path) {
    if ($Path -eq "/") { $Path = "" }
    $Headers = @{
        Authorization = "Bearer $global:accessToken"
        "Content-Type" = "application/json"
    }
    $Body = @{
        path              = $Path
        recursive         = $false
        include_media_info = $false
        include_deleted   = $false
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

# Función para obtener metadata de un elemento (archivo o carpeta)
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

# Función para descargar un archivo desde Dropbox (a una ruta local dada)
function Download-DropboxFile($FilePath, $LocalPath) {
    $Headers = @{ Authorization = "Bearer $global:accessToken" }
    $Headers["Dropbox-API-Arg"] = (@{ path = $FilePath } | ConvertTo-Json -Compress)
    $Url = "https://content.dropboxapi.com/2/files/download"
    
    Debug-Print "Download-DropboxFile - FilePath: '$FilePath'"
    Debug-Print "Dropbox-API-Arg: $($Headers['Dropbox-API-Arg'])"
    
    $FileName = [System.IO.Path]::GetFileName($FilePath)
    if (-not $FileName) {
        [System.Windows.Forms.MessageBox]::Show("Error: Nombre de archivo no válido.")
        return
    }
    $OutFilePath = Join-Path -Path $LocalPath -ChildPath $FileName
    
    try {
        # Se fuerza un body nulo y ContentType vacío para evitar error 400
        Invoke-WebRequest -Uri $Url -Method Post -Headers $Headers -Body $null -ContentType "" -OutFile $OutFilePath -Verbose
        Debug-Print "Archivo descargado en: $OutFilePath"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al descargar archivo: $($_.Exception.Message)")
        Debug-Print "Error al descargar archivo: $($_.Exception.Message)"
    }
}

# Función para descargar una carpeta de Dropbox de manera recursiva
function Download-DropboxFolder($FolderPath, $LocalParent) {
    Debug-Print "Download-DropboxFolder - FolderPath: '$FolderPath'"
    # Verificar que la metadata indica que es carpeta
    $meta = Get-DropboxMetadata $FolderPath
    if (-not $meta -or $meta['.tag'] -ne "folder") {
        Debug-Print "La ruta $FolderPath no es una carpeta válida. Abortando descarga de carpeta."
        return
    }
    
    $folderName = [System.IO.Path]::GetFileName($FolderPath)
    if (-not $folderName) { $folderName = "RootFolder" }
    
    $localFolderPath = Join-Path $LocalParent $folderName
    if (!(Test-Path $localFolderPath)) {
        New-Item -ItemType Directory -Path $localFolderPath | Out-Null
    }
    
    $entries = Get-DropboxFiles -Path $FolderPath
    foreach ($entry in $entries) {
        if ($entry[".tag"] -eq "folder") {
            Download-DropboxFolder $entry.path_lower $localFolderPath
        } else {
            Download-DropboxFile $entry.path_lower $localFolderPath
        }
    }
}

############################
# Crear interfaz gráfica
############################

$form = New-Object System.Windows.Forms.Form
$form.Text = "UTP"
$form.Size = New-Object System.Drawing.Size(700, 500)
$form.StartPosition = "CenterScreen"

# Lista de archivos y carpetas (ventana izquierda)
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 10)
$listBox.Size = New-Object System.Drawing.Size(300, 400)
$form.Controls.Add($listBox)

# Lista de elementos seleccionados (ventana derecha)
$selectedFilesBox = New-Object System.Windows.Forms.ListBox
$selectedFilesBox.Location = New-Object System.Drawing.Point(350, 10)
$selectedFilesBox.Size = New-Object System.Drawing.Size(300, 350)
$form.Controls.Add($selectedFilesBox)

# Función auxiliar para dar estilo a los botones
function Set-ButtonStyle($button) {
    $button.BackColor  = [System.Drawing.Color]::Red
    $button.ForeColor  = [System.Drawing.Color]::White
    $button.FlatStyle  = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $button.FlatAppearance.BorderSize  = 1
    $button.TextAlign  = [System.Drawing.ContentAlignment]::MiddleCenter
}

############################
# Botones de agregar/eliminar (entre las listas)
############################

# Botón agregar (flecha derecha)
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

# Botón eliminar (flecha izquierda)
$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = "←"
$removeButton.Size = New-Object System.Drawing.Size(30,30)
$removeButton.Location = New-Object System.Drawing.Point(320, 240)
Set-ButtonStyle $removeButton

$removeButton.Add_Click({
    $selectedFilesBox.Items.Remove($selectedFilesBox.SelectedItem)
})
$form.Controls.Add($removeButton)

############################
# Botones de subir/bajar (a la derecha de la lista de seleccionados)
############################

# Botón subir (flecha arriba)
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

# Botón bajar (flecha abajo)
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

############################
# Botón de descargar y ejecutar
############################
$downloadButton= New-Object System.Windows.Forms.Button
$downloadButton.Text="Descargar"
$downloadButton.Size=New-Object System.Drawing.Size(150,30)
$downloadButton.Location= New-Object System.Drawing.Point(350, 350)
Set-ButtonStyle $downloadButton

$downloadButton.Add_Click({
    # Abrir diálogo para seleccionar carpeta de descarga
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $chosenPath = $folderDialog.SelectedPath
        Debug-Print "Carpeta elegida: $chosenPath"
        foreach ($item in $selectedFilesBox.Items) {
            $metadata = Get-DropboxMetadata $item
            if ($metadata -and $metadata['.tag'] -eq "folder") {
                Download-DropboxFolder $item $chosenPath
            } else {
                Download-DropboxFile $item $chosenPath
                $FileName = [System.IO.Path]::GetFileName($item)
                $DownloadedFile = Join-Path $chosenPath $FileName
                if (Test-Path $DownloadedFile) {
                    Debug-Print "Ejecutando archivo: $DownloadedFile"
                    Start-Process -FilePath $DownloadedFile
                } else {
                    Debug-Print "El archivo $DownloadedFile no se encontró tras la descarga."
                }
            }
        }
    }
})
$form.Controls.Add($downloadButton)

# Variables globales para la ruta actual y entradas de Dropbox
$global:currentPath = ""
$global:dropboxEntries = @{ }

# Función para actualizar la lista de archivos en la ventana izquierda
function Update-FileList {
    $listBox.Items.Clear()
    
    if ($global:currentPath -ne "") {
        $listBox.Items.Add("..")
    }
    
    $entries = Get-DropboxFiles -Path $global:currentPath
    if ($entries) {
        $global:dropboxEntries = @{ }
        foreach ($entry in $entries) {
            if ($entry.name) {
                $listBox.Items.Add($entry.name)
                $global:dropboxEntries[$entry.name] = $entry
            }
        }
    }
}

# Función auxiliar para obtener la entrada de Dropbox (insensible a mayúsculas)
function Get-DropboxEntry($name) {
    foreach ($key in $global:dropboxEntries.Keys) {
        if ($key.ToLower() -eq $name.ToLower()) {
            return $global:dropboxEntries[$key]
        }
    }
    return $null
}

# Evento de doble clic para navegar entre carpetas en la ventana izquierda
$listBox.Add_DoubleClick({
    $selectedItem = $listBox.SelectedItem.Trim()
    Debug-Print "Double-clicked item: '$selectedItem'"
    if (-not [string]::IsNullOrEmpty($selectedItem)) {
        if ($selectedItem -eq "..") {
            $global:currentPath = [System.IO.Path]::GetDirectoryName($global:currentPath)
            if (-not $global:currentPath) { $global:currentPath = "" }
        } else {
            $entry = Get-DropboxEntry $selectedItem
            if ($entry -and $entry[".tag"] -eq "folder") {
                if ($global:currentPath -eq "") {
                    $global:currentPath = "/$selectedItem"
                } else {
                    $global:currentPath = "$global:currentPath/$selectedItem"
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("El elemento seleccionado no es una carpeta.")
                
            }
        }
        Update-FileList
    }
})

# Obtener el token y cargar archivos
$global:accessToken = Get-AccessToken
if ($global:accessToken) {
    Update-FileList
    $form.ShowDialog()
} 
