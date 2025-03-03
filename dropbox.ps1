Add-Type -AssemblyName System.Windows.Forms

# Función para obtener el token de acceso
function Get-AccessToken {
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

# Función para obtener la lista de archivos y carpetas de Dropbox
function Get-DropboxFiles($Path) {
    $Headers = @{ Authorization = "Bearer $global:accessToken"; "Content-Type" = "application/json" }
    $Body = @{ path = $Path; recursive = $false; include_media_info = $false; include_deleted = $false } | ConvertTo-Json -Depth 10
    $Url = "https://api.dropboxapi.com/2/files/list_folder"
    
    try {
        $Response = Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -Body $Body
        return $Response.entries
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al obtener archivos: $($_.ErrorDetails.Message)")
        return $null
    }
}

# Función para descargar archivos desde Dropbox
function Download-DropboxFile($FilePath) {
    $FilePath = $FilePath.Trim()
    if (-not $FilePath.StartsWith("/")) {
        $FilePath = "/" + $FilePath
    }
    
    $Headers = @{ Authorization = "Bearer $global:accessToken" }
    $Headers["Dropbox-API-Arg"] = (@{ path = $FilePath } | ConvertTo-Json -Compress)
    $Url = "https://content.dropboxapi.com/2/files/download"
    $OutFilePath = Join-Path $env:USERPROFILE "Downloads" (Split-Path $FilePath -Leaf)
    
    try {
        Invoke-RestMethod -Uri $Url -Method Post -Headers $Headers -OutFile $OutFilePath
        [System.Windows.Forms.MessageBox]::Show("Descarga completada en: $OutFilePath")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al descargar archivo: $($_.Exception.Message)")
    }
}

# Crear la ventana
$form = New-Object System.Windows.Forms.Form
$form.Text = "Explorador de Dropbox"
$form.Size = New-Object System.Drawing.Size(500, 500)
$form.StartPosition = "CenterScreen"

# Lista para mostrar archivos y carpetas
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Dock = "Top"
$listBox.Height = 400
$form.Controls.Add($listBox)

# Botón para descargar archivos
$downloadButton = New-Object System.Windows.Forms.Button
$downloadButton.Text = "Descargar"
$downloadButton.Dock = "Bottom"
$downloadButton.Add_Click({
    $selectedItem = $listBox.SelectedItem
    if ($selectedItem -and $selectedItem -ne "..") {
        $filePath = "$global:currentPath/$selectedItem" -replace "//", "/"
        Download-DropboxFile -FilePath $filePath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Seleccione un archivo válido para descargar.")
    }
})
$form.Controls.Add($downloadButton)

# Variable global para la ruta actual en Dropbox
$global:currentPath = ""

# Función para actualizar la lista de archivos
function Update-FileList {
    $listBox.Items.Clear()
    
    if ($global:currentPath -ne "") {
        $listBox.Items.Add("..")
    }
    
    $entries = Get-DropboxFiles -Path $global:currentPath
    if ($entries) {
        foreach ($entry in $entries) {
            if ($entry.name) {
                $listBox.Items.Add($entry.name)
            }
        }
    }
}

# Evento de doble clic para navegar
$listBox.Add_DoubleClick({
    $selectedItem = $listBox.SelectedItem
    if ($selectedItem -eq "..") {
        $global:currentPath = [System.IO.Path]::GetDirectoryName($global:currentPath) -replace "\\", "/"
        if (-not $global:currentPath) { $global:currentPath = "" }
    } else {
        $global:currentPath = "$global:currentPath/$selectedItem" -replace "//", "/"
    }
    Update-FileList
})

# Obtener el token y cargar archivos
$global:accessToken = Get-AccessToken
if ($global:accessToken) {
    Update-FileList
    $form.ShowDialog()
} else {
    [System.Windows.Forms.MessageBox]::Show("No se pudo obtener el token de acceso.")
}
