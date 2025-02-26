# ---------------------------
# Configuración y variables globales
# ---------------------------
$token = "sl.u.AFi2cc9aoCUkCQEMQK2E600o30S9Brg-CEHQH0n8y8xrwG42Zp07fhXGhunkKqLGdLFYFZNM47AcmBv4rEAj8XazYRRLw4OKbNKylfWfLTAZyleXuSNqHBrEuB5UYFt6i8BMV0LYqSRcb9seCuy730-am3PV71DFCmW86YpIUQbEktrlRWJEDJSke3qfao5h3v15z2S2ZV9nX-lM4GKMQt_b9xZb3TKZX14bx90L1stybPUikaQtgPVhJNhTWY6MuBzFHA7kBReAhSsjgbtMNAOdfOfpPOn9L-HIUYUV7i_KQlSLFPEh2ijbXlwXQlDgkFRbFanYCF3uWmD5StT5kHe4H1kI0VG7txxQ-TLG0CFRWLM12VYjvO7F1z6WIML51zN5g7dkqq2qahN1Rsa0x4VWAqoKUflz-92_gaQ6V7-uYw4mLvWhKN6RJtunkbukO8M8abqSxOUarBPz0TaTgynJ3DrD7mvWrmU1irPCFZAoa5ukjVZAHcLvtSeGPfkD7O2lW4oEJTrY7ewVXiNmCZfq08K7PZxk2cnPCc_q5vvOaUdUBM8zLzPn-FmsEN4kKxvCCHT61zRe4yUCtsk8Uo-1_e6z4E9Mtd6Jmix_XoEBUSdnCH7M6m6qVe410yKolGdViK26EfMQEmmSPM8QyrEBLVzMmaULcSnC4qYmaStqiygiJHbW6AHwIGIshGA5qDNzAVMahBnGIQXUy6xNWHc_Pfs8Fzq9JkmIzGOBXv7_zh_JTiO7v4oLhZ8cZkMrTxmcFhiIhbrAbcQGnUIF2K-fl0PJlx0tq1TDXM0wW2bxKicSMlOxZV4v6J_7hxU4IVt3x-AypVJRm36biiI7YWiy6Ib3Kmjd7_fzot2GARS1pLyFqBDY0rGxYClZIeBfFwjglUn-wUQwo2JglaR7OCL3GqY09TmJLTc-EmjnILloZ6zRjrWFTR-YcQOJcGNdmk46w4PvwemtVaxrUJJqLmFHE2bDjzv2UcoOVNjd9VffkM11VuIiiJprZG5Z9uG3JsB-hK-b8wCgnTWd3WmztnJUErxsCWz4qwuKzxosO2ZWZYo7TP1QwG7LkJJHfOZEZhil7lJL-ugfMcRtJHsSbXe3EHZC6nUccpqxIeqM0xz-Z3R17z7KOqPT1ZLDS8MGPBvxPpYkFRU-V1btPr0KmrVbQ8IZAA5e37TK-nbS-hTdo_mV6gSg9ONbOegLfeHK0-comExdeVQBmlK72nueZKB-HTMD9AAJ5r6lHxsdO8jBsp_BAFRyQ6VTBc2YMcaKnaHoiRPpGlQj8aDhUbwJyC7A2cv9kgfh6poJJUJPXDYEXnZVzjxcd-7mMl9LpKhsvnfHPBuGfEXoeX-rAZq3cwXn"

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
