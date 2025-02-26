Add-Type -AssemblyName System.Windows.Forms

# Configuración y variables globales
$token = "sl.u.AFnQcB96oZhA_qWNVkFkiS6QS-u2NgeVd_ZN72ujPChg1en6CUZFG1z9A5Ldl47c8XClOoXtzoPQ4BXzPugllI7abs1tnqu6od4R-M9emhQzg3VXJoKZ9-0tV5GXyfGupX4lfgTwEamx7lMQOQCXBKuL7Yj24N3WV0j0z9ohkRiJAdrMg-HeSWJuqDojNPpNMgFYfAoaiy6yoqamThsTQPX3J7sORDkw8yPdDJZTbHvp28kHBFYdM36DW7jz0AhU3ak00mS88f5vdarg_mb-gv08GvnbGsYkzXNEzefow2z8sxPVzjusUvLLEoy4_TYRDmzCNxgSSFujkswGiPnjeeHqgd7MqWat4KpDTEkI1y5WQXU8scwExjDR2G22x6krP5u-TcXhmdY4gZf6LKVQ-d4gYF9oLV6hwBO2hJciAKjCJTHFL42TBBG96eTR5dyxER2yy-y3dvaCNbYQL4x-A-_IuR5LsMMVVnASeLzmyEY5DJ7a-gOAjIUojr2V5fOuRRD56NcCWOlaDyrAirb9Dca92mSRcUeYct4uSJuNdd26YD6Xm6QquZ4lpDahAU9coRx8Vmf-DociZZcDJfd0XvaH8tHBB4q4QyRThVCH8V9LT2hfuUF7MbAuZS5tVAGzg2w-3aB1IIrCKxAkS1DcMs0oG_47Ndx-ELMaFpy5FUruQzObUQGQvWO4-LiilLFo4AmwltiQBGjgB5V-yR5dpHhM9KBQ2J9tdl4FQefCU1vRpNfPYgI-PTaHg8gAfAuVmq2eKeX-iZeXcPrF1PJCcEEkGdSjIbeOp2odP2K7guapxpaPIIDWWB9PDGFTNMC5BtSRy8VWrGe0wmNDNmnu2fw8jA9M7Vq4JinYXXksxUsG-XkGvZPlDGKDqh4KMMJIGwO3rEAKsKa5rhiIbGvd6UrIGLY1EgMPrpgRvijRWJyWcCyZaF6t5qzHe716JYke4w1M02IhAEEm8ymUNz5nEzOzQ78baLFneK1m1dnyyrVCsmsmPsSFwipu-kHmAUUSg11JeCBP1slHOaVGF32g4Skw93R4YzxDluEe0KLMbWkex1YuWaphjr8JGxkmKfkR1Jt_9yOhSCqxOueEm0gDpBKXNErRgB8j8lSrl1YUgk6UI38FAKMCcZigeTvMwUtoTAjZQ-9K4sWkkQEwep1VI7x9vIhu_2sMI3LBfC5Sj_r6FLfy-4I15eFz6ocrzhg515YShXWlFxag4BifmxlFS6eCloPsEZdtRGa7SVppuPc_am8ILsyhb8Nhdq-sFh0ji3SRgPcnWx7AEdAsocPdpIAxmWWJKTLo2_3xQ01gxCMsicIW_oCHe8jbrPbb3imZ5Un3anDpWAQlQERipdfOWxxT"
$headers = @{ "Authorization" = "Bearer $token"; "Content-Type" = "application/json" }
$global:downloadDirectory = ""
$global:currentPath = ""
$global:history = @()
$global:downloadedFiles = @()

# Funciones para obtener carpetas y archivos
function Get-Folders {
    $url = "https://api.dropboxapi.com/2/files/list_folder"
    $body = @{ path = $global:currentPath } | ConvertTo-Json -Compress
    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        return $response.entries | Where-Object { $_.".tag" -eq "folder" }
    } catch {
        return @()
    }
}

function Get-Files {
    $url = "https://api.dropboxapi.com/2/files/list_folder"
    $body = @{ path = $global:currentPath } | ConvertTo-Json -Compress
    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        return $response.entries | Where-Object { $_.".tag" -eq "file" }
    } catch {
        return @()
    }
}

# Funciones para actualizar listas y descargar archivos
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

function DownloadFiles {
    if ([string]::IsNullOrEmpty($global:downloadDirectory)) {
        [System.Windows.Forms.MessageBox]::Show("Selecciona primero la carpeta de destino.")
        return
    }
    
    $selectedFiles = $listBoxFiles.SelectedItems
    if ($selectedFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No se ha seleccionado ningún archivo.")
        return
    }
    
    $global:downloadedFiles = @()
    
    foreach ($file in $selectedFiles) {
        $dropboxFilePath = if ($global:currentPath -eq "") { "/$file" } else { "$global:currentPath/$file" }
        $localFilePath = Join-Path $global:downloadDirectory $file
        $downloadArgs = @{ path = $dropboxFilePath } | ConvertTo-Json -Compress
        $url = "https://content.dropboxapi.com/2/files/download"
        
        try {
            $listBoxStatus.Items.Add("Descargando: $file")
            $client = New-Object System.Net.WebClient
            $client.Headers.Add("Authorization", "Bearer $token")
            $client.Headers.Add("Dropbox-API-Arg", $downloadArgs)
            $client.DownloadFile($url, $localFilePath)
            $global:downloadedFiles += $localFilePath
            $listBoxStatus.Items.Add("Completado: $file")
        } catch {
            $listBoxStatus.Items.Add("❌ Error al descargar $file")
        }
    }
    
    [System.Windows.Forms.MessageBox]::Show("Descarga finalizada.")
    Open-SortWindow
}

# Función para abrir la ventana de ordenamiento
function Open-SortWindow {
    $sortForm = New-Object System.Windows.Forms.Form
    $sortForm.Text = "Ordenar Archivos"
    $sortForm.Size = New-Object System.Drawing.Size(400, 500)

    $listBoxSort = New-Object System.Windows.Forms.ListBox
    $listBoxSort.Location = New-Object System.Drawing.Point(10, 10)
    $listBoxSort.Size = New-Object System.Drawing.Size(360, 250)
    $listBoxSort.SelectionMode = "One"
    foreach ($file in $global:downloadedFiles) {
        $listBoxSort.Items.Add((Get-Item $file).Name)
    }
    $sortForm.Controls.Add($listBoxSort)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(10, 270)
    $statusLabel.Size = New-Object System.Drawing.Size(360, 100)
    $sortForm.Controls.Add($statusLabel)

    $moveUpButton = New-Object System.Windows.Forms.Button
    $moveUpButton.Text = "Subir"
    $moveUpButton.Location = New-Object System.Drawing.Point(50, 380)
    $moveUpButton.Size = New-Object System.Drawing.Size(100, 30)
    $moveUpButton.Add_Click({ Move-ItemUp $listBoxSort })
    $sortForm.Controls.Add($moveUpButton)

    $moveDownButton = New-Object System.Windows.Forms.Button
    $moveDownButton.Text = "Bajar"
    $moveDownButton.Location = New-Object System.Drawing.Point(160, 380)
    $moveDownButton.Size = New-Object System.Drawing.Size(100, 30)
    $moveDownButton.Add_Click({ Move-ItemDown $listBoxSort })
    $sortForm.Controls.Add($moveDownButton)

    $executeButton = New-Object System.Windows.Forms.Button
    $executeButton.Text = "Ejecutar"
    $executeButton.Location = New-Object System.Drawing.Point(270, 380)
    $executeButton.Size = New-Object System.Drawing.Size(100, 30)
    $executeButton.Add_Click({ Execute-Files ($listBoxSort.Items) $statusLabel })
    $sortForm.Controls.Add($executeButton)

    $sortForm.ShowDialog()
}

# Función para ejecutar archivos en orden y mostrar estado
function Execute-Files {
    param ($orderedFiles, $statusLabel)
    $statusLabel.Text = ""
    
    foreach ($file in $orderedFiles) {
        $fullPath = Join-Path $global:downloadDirectory $file
        if (Test-Path $fullPath) {
            try {
                Start-Process -FilePath $fullPath -Wait
                $statusLabel.Text += "Instalado: $file`n"
            } catch {
                $statusLabel.Text += "No instalado: $file`n"
            }
        } else {
            $statusLabel.Text += "No instalado: $file`n"
        }
    }
}

Update-Lists
$form.ShowDialog()
