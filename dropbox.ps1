# ---------------------------
# Configuración y variables globales
# ---------------------------
$token = "sl.u.AFlq_I0lMDoYN6BPivPAC6ItQUvEQDmFrPFZ6-elVzifAljpG8L4PXQBzSu2wbrFfnT6inppfAi2TpPKDbjvUnUDHb0ohda6Bpu2aSkVFnKqvNAm-n-CXKOWGdjLIn6qZ3n3DK99Oa3-if8CbefPMURagDSVRyPhjs3PglpUDGyyXhIVE764V3pS0nmhPILJMgwZ1BI2I6McNjdgWdiw_dHjnWhbEK2LdAnfEGmMLzHGWt6ynI24kXwAlkqscVi3HH0yjBtXCYYgd9q1_ucC03dpMsLMjpeks7XhGy_ZvpR1WTb3OYMLkUMzxJjDP-kXtA37TZq4ptT0Az-sHVH7GWm-Or64m0Luc2EV_nosE9nE-07YWthzoeYIwRCIAr-Q-nEWIag2RrSPDbrpC4FluOQMGKv1DzBj2INWuwnbJRugS7zGhXczKZxCHBzgsL0wzxKaUdC-cfcv3v2Zo2UNkLEzUuDMkiHcaikAq_KHGll47iDjBA55ZLDD6m8U4K95zOrablVzK_-YJkuhB5ddj_hg6TSDAFNTjMjkx3o2-vm9kY2BZ3Z1tspbNVRZWphfHbGmn4kCKoaR6UN8Pnuo_E9mV3vTTkRVAgrQ3mV3_TGBVWVfij1BzoIkqqSlXdIOS4txdVY1H0YDpohvbWwa59OAzrUbcHoKV_gg5EDxUS5pvPiGKSopuTWMsunE59ZZgCN_vZWivWCWcOEFsJC_lFsIAHHPYCYpxhzfTVid4LwOq5H2kJu7RVx_h4iMY0cwIKlTgdJQ9ZvU_KrwUAbYVyJMrI-hHgit5_OO4RkmOxA7bDcxldk4jXoSsK-sebSlpmoKLZi7XU9g2_5E0qg70L1raMi1eNITL_2zvLPg5i7xy4O1BsCMfG_8bUn0dJlCBtKgmEkkTZI02uIUyWQMiSVcYnBXZt-lVOAI0_vVi8zBZax-_B0jzHo0J2vgECCAfxUjyKL6DJyni7ieHtRblLZM713OiJE-AkfcgA0OKzJJjWSwdgEO1LJ9xCdoEWToIGs2uNCRCb1qaT8MG_pyGVWEPTT2UVjxZ2J1b0bT8niHwxkinVlqXCEQAqxpPQz2DWAj5CFQtGG02AAw0Gni40pxnY9VpI0aVtJg34MibjsP810O4oDvEqxhC7ZPyge9gc0jD5wAwIN4w4JhRgjc8bwGNtw0XeMKjyKNWw8ScwhvMNQRDswheaj9oGXLO8R_eTW7crdDT8x2hqObETHg_VYdBHNvOpkoOMalcQS9CWsnLFfb-bQ6HkvYGx83RNxPDiuaeHK8pIHzkLPLLyLViSNYCOtCSGwuWEDAfU58bGK0YYcFThuHxRIIP6QMNSHf40ODRY8Pq4C_7tlqVyD0V3CA"

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type"  = "application/json"
}

$global:downloadDirectory = ""
$global:currentPath = ""
$global:history = @()
$global:cumulativeBytes = 0
$global:totalBytes = 0
$global:startTime = $null
$global:downloadedFiles = @()

# ---------------------------
# Función de navegación: RefreshFolder
# ---------------------------
function RefreshFolder($folderPath) {
    $global:currentPath = $folderPath
    if ($folderPath -eq "") {
        $currentFolderLabel.Text = "Carpeta actual: Raíz"
    } else {
        $currentFolderLabel.Text = "Carpeta actual: " + $folderPath
    }
    $listBoxFolders.Items.Clear()
    $listBoxFiles.Items.Clear()
    $body = '{ "path": "' + $folderPath + '" }'
    try {
        $response = Invoke-RestMethod -Uri "https://api.dropboxapi.com/2/files/list_folder" -Method Post -Headers $headers -Body $body
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al conectar con Dropbox: " + $_.Exception.Message)
        return
    }
    $folders = $response.entries | Where-Object { $_.PSObject.Properties[".tag"].Value -eq "folder" }
    foreach ($folder in $folders) {
         $listBoxFolders.Items.Add($folder.name)
    }
    $files = $response.entries | Where-Object { $_.PSObject.Properties[".tag"].Value -eq "file" }
    foreach ($file in $files) {
         $listBoxFiles.Items.Add($file)
         Write-Host "Archivo: $($file.name) - Tamaño: $($file.size)"
    }
    $listBoxFiles.DisplayMember = "name"
}

# ---------------------------
# Función asíncrona para descargar archivos usando Task.Run
# ---------------------------
function DownloadFilesAsync($selectedFiles) {
    $global:downloadedFiles = @()
    [System.Threading.Tasks.Task]::Run([System.Action]{
        foreach ($file in $selectedFiles) {
            if ($global:currentPath -eq "") {
                $dropboxFilePath = "/" + $file.name
            } else {
                $dropboxFilePath = $global:currentPath + "/" + $file.name
            }
            Write-Host "Intentando descargar: $dropboxFilePath"
            $localFilePath = Join-Path $global:downloadDirectory $file.name
            $downloadArgs = '{"path": "' + $dropboxFilePath + '"}'
            Write-Host "Descargando con argumentos: $downloadArgs"
            $url = "https://content.dropboxapi.com/2/files/download"
            try {
                $request = [System.Net.HttpWebRequest]::Create($url)
                $request.Method = "POST"
                $request.Headers.Add("Authorization", "Bearer $token")
                $request.Headers.Add("Dropbox-API-Arg", $downloadArgs)
                $response = $request.GetResponse()
                $stream = $response.GetResponseStream()
                $fileStream = [System.IO.File]::Create($localFilePath)
                $buffer = New-Object byte[] 8192
                while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $fileStream.Write($buffer, 0, $read)
                    $global:cumulativeBytes += $read
                    $percent = [math]::Round(($global:cumulativeBytes / $global:totalBytes) * 100)
                    $form.Invoke({
                        $progressBar.Value = $percent
                        $progressLabel.Text = "Progreso: $percent %"
                        $elapsed = (Get-Date) - $global:startTime
                        if ($elapsed.TotalSeconds -gt 0) {
                            $speed = $global:cumulativeBytes / $elapsed.TotalSeconds
                            $remainingBytes = $global:totalBytes - $global:cumulativeBytes
                            $estimatedSeconds = $remainingBytes / $speed
                            $timeLabel.Text = "Tiempo restante estimado: " + [Math]::Round($estimatedSeconds,1) + " seg"
                        }
                    })
                    Write-Host "Bytes leídos: $global:cumulativeBytes de $global:totalBytes ($percent %)"
                }
                $fileStream.Close()
                $stream.Close()
                $response.Close()
                $global:downloadedFiles += $localFilePath
            } catch {
                Write-Host "Error al descargar $($dropboxFilePath): $($_.Exception.Message)"
            }
        }
        $form.Invoke({
            ShowExecutionOrderForm $global:downloadedFiles
        })
    })
}

# ---------------------------
# Función para mostrar el formulario de orden de ejecución
# ---------------------------
function ShowExecutionOrderForm($files) {
    $orderForm = New-Object System.Windows.Forms.Form
    $orderForm.Text = "Orden de ejecución"
    $orderForm.Size = New-Object System.Drawing.Size(400,400)
    $orderForm.StartPosition = "CenterScreen"

    $listBoxOrder = New-Object System.Windows.Forms.ListBox
    $listBoxOrder.Location = New-Object System.Drawing.Point(10,10)
    $listBoxOrder.Size = New-Object System.Drawing.Size(360,300)
    foreach ($file in $files) {
        $listBoxOrder.Items.Add($file)
    }
    $orderForm.Controls.Add($listBoxOrder)

    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Text = "Subir"
    $btnUp.Location = New-Object System.Drawing.Point(10,320)
    $orderForm.Controls.Add($btnUp)

    $btnDown = New-Object System.Windows.Forms.Button
    $btnDown.Text = "Bajar"
    $btnDown.Location = New-Object System.Drawing.Point(100,320)
    $orderForm.Controls.Add($btnDown)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Ejecutar"
    $btnRun.Location = New-Object System.Drawing.Point(190,320)
    $orderForm.Controls.Add($btnRun)

    $btnUp.Add_Click({
         if($listBoxOrder.SelectedIndex -gt 0) {
             $index = $listBoxOrder.SelectedIndex
             $item = $listBoxOrder.SelectedItem
             $listBoxOrder.Items.RemoveAt($index)
             $listBoxOrder.Items.Insert($index - 1, $item)
             $listBoxOrder.SelectedIndex = $index - 1
         }
    })

    $btnDown.Add_Click({
         if($listBoxOrder.SelectedIndex -lt $listBoxOrder.Items.Count - 1) {
             $index = $listBoxOrder.SelectedIndex
             $item = $listBoxOrder.SelectedItem
             $listBoxOrder.Items.RemoveAt($index)
             $listBoxOrder.Items.Insert($index + 1, $item)
             $listBoxOrder.SelectedIndex = $index + 1
         }
    })

    $btnRun.Add_Click({
         foreach ($item in $listBoxOrder.Items) {
              Write-Host "Ejecutando: $item"
              Start-Process $item
         }
         [System.Windows.Forms.MessageBox]::Show("Se han ejecutado los programas en el orden especificado.")
         $orderForm.Close()
    })

    $orderForm.ShowDialog()
}

# ---------------------------
# Configuración de la interfaz gráfica principal
# ---------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Navegador de Dropbox"
$form.Size = New-Object System.Drawing.Size(800,680)
$form.StartPosition = "CenterScreen"

$currentFolderLabel = New-Object System.Windows.Forms.Label
$currentFolderLabel.Location = New-Object System.Drawing.Point(10,10)
$currentFolderLabel.Size = New-Object System.Drawing.Size(600,20)
$currentFolderLabel.Text = "Carpeta actual: Raíz"
$form.Controls.Add($currentFolderLabel)

$destinationLabel = New-Object System.Windows.Forms.Label
$destinationLabel.Location = New-Object System.Drawing.Point(10,580)
$destinationLabel.Size = New-Object System.Drawing.Size(600,20)
$destinationLabel.Text = "Destino de descarga: Ninguno"
$form.Controls.Add($destinationLabel)

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(10,550)
$progressLabel.Size = New-Object System.Drawing.Size(200,20)
$progressLabel.Text = "Progreso: 0 %"
$form.Controls.Add($progressLabel)

$timeLabel = New-Object System.Windows.Forms.Label
$timeLabel.Location = New-Object System.Drawing.Point(220,550)
$timeLabel.Size = New-Object System.Drawing.Size(200,20)
$timeLabel.Text = "Tiempo restante estimado: -"
$form.Controls.Add($timeLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,520)
$progressBar.Size = New-Object System.Drawing.Size(760,30)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$form.Controls.Add($progressBar)

$listBoxFolders = New-Object System.Windows.Forms.ListBox
$listBoxFolders.Location = New-Object System.Drawing.Point(10,40)
$listBoxFolders.Size = New-Object System.Drawing.Size(350,350)
$listBoxFolders.Font = New-Object System.Drawing.Font("Segoe UI",10)
$form.Controls.Add($listBoxFolders)

$listBoxFiles = New-Object System.Windows.Forms.ListBox
$listBoxFiles.Location = New-Object System.Drawing.Point(370,40)
$listBoxFiles.Size = New-Object System.Drawing.Size(400,350)
$listBoxFiles.Font = New-Object System.Drawing.Font("Segoe UI",10)
$listBoxFiles.SelectionMode = "MultiExtended"
$listBoxFiles.DisplayMember = "name"
$form.Controls.Add($listBoxFiles)

$backButton = New-Object System.Windows.Forms.Button
$backButton.Location = New-Object System.Drawing.Point(10,400)
$backButton.Size = New-Object System.Drawing.Size(80,30)
$backButton.Text = "Atrás"
$form.Controls.Add($backButton)

$homeButton = New-Object System.Windows.Forms.Button
$homeButton.Location = New-Object System.Drawing.Point(100,400)
$homeButton.Size = New-Object System.Drawing.Size(80,30)
$homeButton.Text = "Inicio"
$form.Controls.Add($homeButton)

$downloadButton = New-Object System.Windows.Forms.Button
$downloadButton.Location = New-Object System.Drawing.Point(370,400)
$downloadButton.Size = New-Object System.Drawing.Size(100,30)
$downloadButton.Text = "Descargar"
$form.Controls.Add($downloadButton)

$selectDestinationButton = New-Object System.Windows.Forms.Button
$selectDestinationButton.Location = New-Object System.Drawing.Point(480,400)
$selectDestinationButton.Size = New-Object System.Drawing.Size(150,30)
$selectDestinationButton.Text = "Seleccionar destino"
$form.Controls.Add($selectDestinationButton)

# ---------------------------
# Eventos de la interfaz principal
# ---------------------------
$listBoxFolders.Add_DoubleClick({
    if ($listBoxFolders.SelectedItem -ne $null) {
        $selectedFolder = $listBoxFolders.SelectedItem.ToString()
        if ($global:currentPath -eq "") {
            $newPath = "/" + $selectedFolder
        } else {
            $newPath = $global:currentPath + "/" + $selectedFolder
        }
        $global:history += $global:currentPath
        RefreshFolder $newPath
    }
})

$backButton.Add_Click({
    if ($global:history.Count -gt 0) {
         $previousPath = $global:history[-1]
         $global:history = $global:history[0..($global:history.Count - 2)]
         RefreshFolder $previousPath
    } else {
         [System.Windows.Forms.MessageBox]::Show("No hay carpeta anterior.")
    }
})

$homeButton.Add_Click({
    $global:history = @()
    RefreshFolder ""
})

$downloadButton.Add_Click({
    if ([string]::IsNullOrEmpty($global:downloadDirectory)) {
         [System.Windows.Forms.MessageBox]::Show("Selecciona primero la carpeta de destino para la descarga.")
         return
    }
    $selectedFiles = $listBoxFiles.SelectedItems
    if ($selectedFiles.Count -eq 0) {
         [System.Windows.Forms.MessageBox]::Show("No se ha seleccionado ningún archivo.")
         return
    }
    $global:totalBytes = 0
    foreach ($file in $selectedFiles) {
         if ($file.size -ne $null) {
              $global:totalBytes += [int64]$file.size
         }
    }
    if ($global:totalBytes -eq 0) {
         [System.Windows.Forms.MessageBox]::Show("No se pudo determinar el tamaño total de descarga.")
         return
    }
    Write-Host "Total bytes a descargar: $global:totalBytes"
    $progressBar.Value = 0
    $global:cumulativeBytes = 0
    $global:startTime = Get-Date
    DownloadFilesAsync $selectedFiles
})

$selectDestinationButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Selecciona la carpeta de destino para la descarga"
    $folderBrowser.ShowNewFolderButton = $true
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
         $global:downloadDirectory = $folderBrowser.SelectedPath
         $destinationLabel.Text = "Destino de descarga: " + $global:downloadDirectory
    }
})

RefreshFolder ""
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
