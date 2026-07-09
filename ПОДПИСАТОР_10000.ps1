# Подключаем необходимые библиотеки
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Xml

# --- НАСТРОЙКИ ---
$CryptCPPath = "C:\Program Files\Crypto Pro\CSP\cryptcp.exe"
if (!(Test-Path $CryptCPPath)) {
    $CryptCPPath = "C:\Program Files (x86)\Crypto Pro\CSP\cryptcp.exe"
}
if (!(Test-Path $CryptCPPath)) {
    [System.Windows.Forms.MessageBox]::Show("Ошибка: Не найдена утилита cryptcp.exe", "Ошибка", "OK", "Error")
    exit
}

# --- ФОРМА ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Подписание XML/GGE (KEП) - v4.0"
$form.Size = New-Object System.Drawing.Size(920, 920)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(10, 10)
$panel.Size = New-Object System.Drawing.Size(880, 890)
$panel.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($panel)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Авто-подписание XML (Универсальный парсер)"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(20, 15)
$lblTitle.Size = New-Object System.Drawing.Size(840, 30)
$lblTitle.TextAlign = "MiddleCenter"
$panel.Controls.Add($lblTitle)

# --- DRAG & DROP ЗОНА ---
$dropZone = New-Object System.Windows.Forms.Panel
$dropZone.Location = New-Object System.Drawing.Point(20, 60)
$dropZone.Size = New-Object System.Drawing.Size(840, 110)
$dropZone.BackColor = [System.Drawing.Color]::FromArgb(230, 240, 255)
$dropZone.BorderStyle = "FixedSingle"
$dropZone.AllowDrop = $true
$panel.Controls.Add($dropZone)

$lblDrop = New-Object System.Windows.Forms.Label
$lblDrop.Text = "Перетащите XML/GGE файлы сюда или кликните для выбора"
$lblDrop.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$lblDrop.ForeColor = [System.Drawing.Color]::DarkBlue
$lblDrop.TextAlign = "MiddleCenter"
$lblDrop.Dock = "Fill"
$lblDrop.BackColor = [System.Drawing.Color]::Transparent
$lblDrop.Cursor = "Hand"
$dropZone.Controls.Add($lblDrop)

$lblFileCount = New-Object System.Windows.Forms.Label
$lblFileCount.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblFileCount.ForeColor = [System.Drawing.Color]::DarkGreen
$lblFileCount.Location = New-Object System.Drawing.Point(20, 175)
$lblFileCount.Size = New-Object System.Drawing.Size(840, 18)
$lblFileCount.TextAlign = "MiddleCenter"
$panel.Controls.Add($lblFileCount)

# ТАБЛИЦА
$gridSignatures = New-Object System.Windows.Forms.DataGridView
$gridSignatures.Location = New-Object System.Drawing.Point(20, 200)
$gridSignatures.Size = New-Object System.Drawing.Size(840, 180)
$gridSignatures.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$gridSignatures.BorderStyle = "FixedSingle"
$gridSignatures.AllowUserToAddRows = $false
$gridSignatures.AllowUserToDeleteRows = $false
$gridSignatures.ReadOnly = $true
$gridSignatures.AutoSizeColumnsMode = "Fill"
$gridSignatures.ColumnCount = 4
$gridSignatures.Columns[0].Name = "Файл"
$gridSignatures.Columns[1].Name = "Роль"
$gridSignatures.Columns[2].Name = "ФИО (из XML)"
$gridSignatures.Columns[3].Name = "Сертификат"
$gridSignatures.Columns[3].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Green
$panel.Controls.Add($gridSignatures)

# Переменные
$script:SelectedFilePaths = @()
$script:AllCertificates = @()
$script:CheckedThumbprints = @()
$script:FileSignatures = @{} # Данные из XML: Файл -> Список подписантов
$script:FileCertMap = @{}    # Карта подписания: Файл -> Список отпечатков

# --- ЛОГИКА DRAG & DROP (Исправлено для 8+ файлов) ---
$dropZone.Add_DragEnter({
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        $dropZone.BackColor = [System.Drawing.Color]::FromArgb(200, 230, 255)
    }
})
$dropZone.Add_DragLeave({ $dropZone.BackColor = [System.Drawing.Color]::FromArgb(230, 240, 255) })
$dropZone.Add_DragDrop({
    [string[]]$files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -gt 0) {
        $script:SelectedFilePaths = $files
        UpdateFileLabel
    }
    $dropZone.BackColor = [System.Drawing.Color]::FromArgb(230, 240, 255)
})

$ClickAction = {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Выберите XML/GGE файлы"
    $dialog.Filter = "XML файлы|*.xml;*.gge|Все файлы|*.*"
    $dialog.Multiselect = $true
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:SelectedFilePaths = [string[]]$dialog.FileNames
        UpdateFileLabel
    }
}
$dropZone.Add_Click($ClickAction)
$lblDrop.Add_Click($ClickAction)

function UpdateFileLabel {
    if ($script:SelectedFilePaths.Count -gt 0) {
        $lblDrop.Text = "Файлов выбрано: $($script:SelectedFilePaths.Count)"
        $lblFileCount.Text = ($script:SelectedFilePaths | ForEach-Object { Split-Path $_ -Leaf }) -join " | "
        ReadSignaturesFromXML
        LoadCertificates
    }
}

# --- УНИВЕРСАЛЬНЫЙ ПАРСИНГ XML ---
function ExtractSurname {
    param([string]$fullName)
    $clean = $fullName.Trim() -replace '\s+', ' '
    $parts = $clean -split ' '
    if ($parts.Count -eq 0) { return $clean }
    
    # 1. Ищем часть без точек (Фамилия)
    $surname = $parts | Where-Object { $_ -notmatch '\.' } | Sort-Object Length -Descending | Select-Object -First 1
    if ($surname) { return $surname }
    
    # 2. Если всё с точками, берем самую длинную часть
    return ($parts | Sort-Object Length -Descending | Select-Object -First 1)
}

function ReadSignaturesFromXML {
    $script:FileSignatures = @{}
    $script:FileCertMap = @{}
    $gridSignatures.Rows.Clear()
    
    $processedCount = 0
    $noSignaturesCount = 0
    $errorCount = 0
    
    foreach ($filePath in $script:SelectedFilePaths) {
        try {
            $xml = [Xml.XmlDocument]::new()
            $xml.Load($filePath)
            $fileName = Split-Path $filePath -Leaf
            $signatures = @()
            
            $signaturesNode = $xml.SelectSingleNode("//Signatures")
            if ($signaturesNode) {
                # Проходим по всем ролям (Composer, Verifier, Chief и т.д.)
                foreach ($roleNode in $signaturesNode.ChildNodes | Where-Object { $_.NodeType -eq 'Element' }) {
                    $fullName = ""
                    
                    # СТРАТЕГИЯ ПОИСКА ИМЕНИ:
                    # 1. Пытаемся найти тег <Name> внутри роли (игнорирует <Department>, <Position>)
                    $nameNode = $roleNode.SelectSingleNode("Name")
                    if ($nameNode) {
                        $fullName = $nameNode.InnerText.Trim()
                    } 
                    # 2. Если тега <Name> нет, берем просто текст из роли (для <ComposeFIO>)
                    elseif ($roleNode.InnerText.Trim()) {
                        $fullName = $roleNode.InnerText.Trim()
                    }

                    if ($fullName -and $fullName -ne "") {
                        $surname = ExtractSurname $fullName
                        $roleName = $roleNode.Name
                        
                        # Делаем роли читаемыми
                        switch ($roleName) {
                            "ComposeFIO" { $roleName = "Составитель" }
                            "VerifyFIO"  { $roleName = "Проверяющий" }
                            "Composer"   { $roleName = "Составитель" }
                            "Verifier"   { $roleName = "Проверяющий" }
                        }
                        
                        $signatures += @{
                            File     = $fileName
                            Path     = $filePath
                            Role     = $roleName
                            FullName = $fullName
                            Surname  = $surname
                            Cert     = ""
                        }
                    }
                }
                
                # Дедупликация внутри одного файла
                $uniqueSignatures = @()
                $seenInFile = @{}
                foreach ($sig in $signatures) {
                    $key = $sig.FullName.ToLower().Trim()
                    if (-not $seenInFile.ContainsKey($key)) {
                        $seenInFile[$key] = $true
                        $uniqueSignatures += $sig
                    }
                }

                if ($uniqueSignatures.Count -gt 0) {
                    $script:FileSignatures[$filePath] = $uniqueSignatures
                    $script:FileCertMap[$filePath] = @()
                    
                    foreach ($sig in $uniqueSignatures) {
                        $gridSignatures.Rows.Add($sig.File, $sig.Role, $sig.FullName, "Ожидание...") | Out-Null
                    }
                    $processedCount++
                } else {
                    $noSignaturesCount++
                    $gridSignatures.Rows.Add($fileName, "Нет данных", "Не найдены ФИО", "Пропущено") | Out-Null
                }
            } else {
                $noSignaturesCount++
                $gridSignatures.Rows.Add($fileName, "Ошибка", "Нет узла <Signatures>", "Пропущено") | Out-Null
            }
        } catch {
            $errorCount++
            $gridSignatures.Rows.Add((Split-Path $filePath -Leaf), "Ошибка чтения", $_.Exception.Message, "Пропущено") | Out-Null
        }
    }
    
    # Принудительное обновление UI (фикс проблемы с 2 файлами)
    $gridSignatures.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    
    $totalSigs = ($script:FileSignatures.Values | Measure-Object Count -Sum).Sum
    $lblStatus.Text = "Файлов: $($script:SelectedFilePaths.Count) | Обработано: $processedCount | Всего подписей: $totalSigs"
    
    if ($processedCount -lt $script:SelectedFilePaths.Count) {
        $msg = "ВНИМАНИЕ:`nОбработано только $processedCount из $($script:SelectedFilePaths.Count) файлов.`nПроверьте структуру XML."
        [System.Windows.Forms.MessageBox]::Show($msg, "Информация", "OK", "Information")
    }
}

# --- ПОИСК И СЕРТИФИКАТЫ ---
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Поиск сертификата по ФИО:"
$lblSearch.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSearch.Location = New-Object System.Drawing.Point(20, 395)
$lblSearch.Size = New-Object System.Drawing.Size(200, 20)
$panel.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(230, 395)
$txtSearch.Size = New-Object System.Drawing.Size(450, 20)
$txtSearch.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$panel.Controls.Add($txtSearch)
$txtSearch.Add_TextChanged({ FilterCertificates })

$btnClearSearch = New-Object System.Windows.Forms.Button
$btnClearSearch.Text = "Сброс"
$btnClearSearch.Location = New-Object System.Drawing.Point(690, 395)
$btnClearSearch.Size = New-Object System.Drawing.Size(170, 20)
$btnClearSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$panel.Controls.Add($btnClearSearch)
$btnClearSearch.Add_Click({ $txtSearch.Text = ""; FilterCertificates })

$certList = New-Object System.Windows.Forms.CheckedListBox
$certList.Location = New-Object System.Drawing.Point(20, 425)
$certList.Size = New-Object System.Drawing.Size(840, 140)
$certList.CheckOnClick = $true
$certList.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$panel.Controls.Add($certList)

$lblCertCount = New-Object System.Windows.Forms.Label
$lblCertCount.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCertCount.ForeColor = [System.Drawing.Color]::Gray
$lblCertCount.Location = New-Object System.Drawing.Point(20, 570)
$lblCertCount.Size = New-Object System.Drawing.Size(840, 18)
$lblCertCount.TextAlign = "MiddleRight"
$panel.Controls.Add($lblCertCount)

# КНОПКИ
$btnAutoSelect = New-Object System.Windows.Forms.Button
$btnAutoSelect.Text = "АВТО-ВЫБОР ПО ФАЙЛАМ"
$btnAutoSelect.Location = New-Object System.Drawing.Point(20, 600)
$btnAutoSelect.Size = New-Object System.Drawing.Size(280, 45)
$btnAutoSelect.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnAutoSelect.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
$btnAutoSelect.ForeColor = "White"
$btnAutoSelect.FlatStyle = "Flat"
$btnAutoSelect.FlatAppearance.BorderSize = 0
$panel.Controls.Add($btnAutoSelect)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "СБРОСИТЬ ВСЕ"
$btnClear.Location = New-Object System.Drawing.Point(310, 600)
$btnClear.Size = New-Object System.Drawing.Size(280, 45)
$btnClear.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$btnClear.BackColor = [System.Drawing.Color]::FromArgb(231, 76, 60)
$btnClear.ForeColor = "White"
$btnClear.FlatStyle = "Flat"
$btnClear.FlatAppearance.BorderSize = 0
$panel.Controls.Add($btnClear)

$btnSign = New-Object System.Windows.Forms.Button
$btnSign.Text = "ПОДПИСАТЬ ВЫБРАННЫЕ"
$btnSign.Location = New-Object System.Drawing.Point(600, 600)
$btnSign.Size = New-Object System.Drawing.Size(260, 45)
$btnSign.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$btnSign.BackColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
$btnSign.ForeColor = "White"
$btnSign.FlatStyle = "Flat"
$btnSign.Enabled = $false
$btnSign.FlatAppearance.BorderSize = 0
$panel.Controls.Add($btnSign)

# СТАТУС
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ожидание файлов..."
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$lblStatus.Location = New-Object System.Drawing.Point(20, 660)
$lblStatus.Size = New-Object System.Drawing.Size(840, 25)
$lblStatus.BorderStyle = "FixedSingle"
$lblStatus.TextAlign = "MiddleCenter"
$panel.Controls.Add($lblStatus)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 700)
$progressBar.Size = New-Object System.Drawing.Size(840, 20)
$panel.Controls.Add($progressBar)

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblProgress.Location = New-Object System.Drawing.Point(20, 730)
$lblProgress.Size = New-Object System.Drawing.Size(840, 20)
$lblProgress.TextAlign = "MiddleCenter"
$panel.Controls.Add($lblProgress)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Font = New-Object System.Drawing.Font("Consolas", 8)
$lblLog.ForeColor = "Red"
$lblLog.BackColor = "#FFF5F5"
$lblLog.BorderStyle = "FixedSingle"
$lblLog.Visible = $false
$lblLog.Location = New-Object System.Drawing.Point(20, 760)
$lblLog.Size = New-Object System.Drawing.Size(840, 50)
$panel.Controls.Add($lblLog)

# --- ФУНКЦИИ СЕРТИФИКАТОВ ---
function LoadCertificates {
    $script:AllCertificates = @()
    try {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store "My", "CurrentUser"
        $store.Open("ReadOnly")
        foreach ($cert in $store.Certificates) {
            if ($cert.HasPrivateKey -and $cert.NotAfter -gt (Get-Date)) {
                $sub = $cert.Subject
                $thumb = $cert.Thumbprint -replace ' ', ''
                $sn=""; $g=""; $i=""
                if ($sub -match "(?:SN|S)=([^,]+)") { $sn=$matches[1].Trim() }
                elseif ($sub -match "(?:CN|E)=([^,]+)") { $sn=$matches[1].Trim() }
                if ($sub -match "(?:G|GN)=([^,]+)") { $g=$matches[1].Trim() }
                if ($sub -match "(?:I|IN)=([^,]+)") { $i=$matches[1].Trim() }
                $clean = "$sn $g $i".Trim()
                if (-not $clean) { $clean = "Сертификат_$($thumb.Substring(0,8))" }
                $script:AllCertificates += @{ Thumbprint=$thumb; CleanName=$clean; Surname=$sn }
            }
        }
        $store.Close()
        $script:AllCertificates = $script:AllCertificates | Sort-Object CleanName
        if ($script:AllCertificates.Count -eq 0) {
            $certList.Items.Clear(); $certList.Items.Add("Нет сертификатов с закрытым ключом!")
            $btnSign.Enabled = $false
        } else {
            $btnSign.Enabled = $true
            FilterCertificates
        }
    } catch { $lblStatus.Text = "Ошибка загрузки сертификатов: $_" }
}

function FilterCertificates {
    $currentlyChecked = @()
    for ($i=0; $i -lt $certList.Items.Count; $i++) {
        if ($certList.GetItemChecked($i)) {
            $d = $certList.Items[$i].ToString()
            if ($d -match '\(([A-F0-9]{8})\.\.\.\)') { $currentlyChecked += $script:AllCertificates | Where-Object { $_.Thumbprint.Substring(0,8) -eq $matches[1] } | Select-Object -ExpandProperty Thumbprint -First 1 }
        }
    }
    foreach ($t in $currentlyChecked) { if ($script:CheckedThumbprints -notcontains $t) { $script:CheckedThumbprints += $t } }

    $certList.Items.Clear()
    $script:CertThumbprints = @()
    $script:CertCleanNames = @()
    $search = $txtSearch.Text.Trim().ToLower()
    foreach ($c in $script:AllCertificates) {
        if (-not $search -or $c.CleanName.ToLower().Contains($search) -or $c.Surname.ToLower().Contains($search)) {
            $certList.Items.Add("$($c.CleanName) ($($c.Thumbprint.Substring(0,8))...)")
            $script:CertThumbprints += $c.Thumbprint
            $script:CertCleanNames += $c.CleanName
        }
    }
    for ($i=0; $i -lt $certList.Items.Count; $i++) {
        if ($script:CheckedThumbprints -contains $script:CertThumbprints[$i]) { $certList.SetItemChecked($i, $true) }
    }
    $lblCertCount.Text = "Показано: $($certList.Items.Count) из $($script:AllCertificates.Count) | Выбрано в UI: $($script:CheckedThumbprints.Count)"
}

function UpdateCertCount { $lblCertCount.Text = "Показано: $($certList.Items.Count) из $($script:AllCertificates.Count) | Выбрано: $($script:CheckedThumbprints.Count)" }

$certList.Add_ItemCheck({
    Start-Sleep -Milliseconds 30
    $idx = $_.Index
    $state = $_.NewValue
    if ($idx -ge 0 -and $idx -lt $script:CertThumbprints.Count) {
        $t = $script:CertThumbprints[$idx]
        if ($state -eq [System.Windows.Forms.CheckState]::Checked) { if ($script:CheckedThumbprints -notcontains $t) { $script:CheckedThumbprints += $t } }
        else { $script:CheckedThumbprints = $script:CheckedThumbprints | Where-Object { $_ -ne $t } }
        UpdateCertCount
    }
})

# --- АВТО-ВЫБОР (ПРИВЯЗКА СЕРТИФИКАТОВ К ФАЙЛАМ) ---
$btnAutoSelect.Add_Click({
    if ($script:FileSignatures.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Нет загруженных файлов!", "Внимание"); return }
    
    $script:FileCertMap = @{}
    $script:CheckedThumbprints = @()
    $notFound = @()

    foreach ($filePath in $script:SelectedFilePaths) {
        $script:FileCertMap[$filePath] = @()
        $sigs = $script:FileSignatures[$filePath]
        if (-not $sigs) { continue }
        
        foreach ($sig in $sigs) {
            # Ищем сертификат по фамилии
            $match = $script:AllCertificates | Where-Object { $_.Surname -eq $sig.Surname } | Select-Object -First 1
            
            if ($match) {
                $sig.Cert = $match.CleanName
                # Добавляем в карту файла (только уникальные для этого файла)
                if ($script:FileCertMap[$filePath] -notcontains $match.Thumbprint) {
                    $script:FileCertMap[$filePath] += $match.Thumbprint
                }
                if ($script:CheckedThumbprints -notcontains $match.Thumbprint) { $script:CheckedThumbprints += $match.Thumbprint }
            } else {
                $sig.Cert = "Не найден"
                $notFound += "$($sig.File) | $($sig.Role): $($sig.FullName) (Фамилия: $($sig.Surname))"
            }
        }
    }

    # Обновляем таблицу
    for ($i=0; $i -lt $gridSignatures.Rows.Count; $i++) {
        $row = $gridSignatures.Rows[$i]
        $fPath = $script:SelectedFilePaths | Where-Object { (Split-Path $_ -Leaf) -eq $row.Cells[0].Value } | Select-Object -First 1
        $role = $row.Cells[1].Value
        $fullName = $row.Cells[2].Value
        
        if ($fPath -and $script:FileSignatures.ContainsKey($fPath)) {
            $foundSig = $script:FileSignatures[$fPath] | Where-Object { $_.Role -eq $role -and $_.FullName -eq $fullName } | Select-Object -First 1
            if ($foundSig) {
                $row.Cells[3].Value = $foundSig.Cert
                $row.Cells[3].Style.ForeColor = if ($foundSig.Cert -eq "Не найден") { "Red" } else { "Green" }
            }
        }
    }
    $gridSignatures.Refresh()
    
    # Обновляем чекбоксы
    for ($i=0; $i -lt $certList.Items.Count; $i++) {
        $certList.SetItemChecked($i, $script:CheckedThumbprints -contains $script:CertThumbprints[$i])
    }
    UpdateCertCount
    $lblStatus.Text = "Авто-выбор выполнен! Файлов: $($script:FileSignatures.Count). Не найдено: $($notFound.Count)"
    if ($notFound.Count -gt 0) { [System.Windows.Forms.MessageBox]::Show("Не найдены сертификаты для:`n" + ($notFound -join "`n"), "Предупреждение", "OK", "Warning") }
})

$btnClear.Add_Click({
    $script:CheckedThumbprints = @()
    $script:FileCertMap = @{}
    for ($i=0; $i -lt $gridSignatures.Rows.Count; $i++) { $gridSignatures.Rows[$i].Cells[3].Value = "Ожидание..."; $gridSignatures.Rows[$i].Cells[3].Style.ForeColor = "Gray" }
    for ($i=0; $i -lt $certList.Items.Count; $i++) { $certList.SetItemChecked($i, $false) }
    $gridSignatures.Refresh()
    UpdateCertCount
    $lblStatus.Text = "Выбор сброшен."
})

# --- ПОДПИСАНИЕ (СТРОГО ПО КАРТЕ ФАЙЛОВ) ---
$btnSign.Add_Click({
    if ($script:SelectedFilePaths.Count -eq 0 -or $script:FileCertMap.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Выберите файлы и выполните авто-выбор!", "Внимание"); return }
    
    $totalOps = 0
    foreach ($f in $script:SelectedFilePaths) { if ($script:FileCertMap[$f]) { $totalOps += $script:FileCertMap[$f].Count } }
    if ($totalOps -eq 0) { [System.Windows.Forms.MessageBox]::Show("Нет совпавших сертификатов для подписания!", "Внимание"); return }

    $res = [System.Windows.Forms.MessageBox]::Show("Будет создано $totalOps подписей.`nКаждый файл будет подписан только своими сертификатами.`nПродолжить?", "Подтверждение", "YesNo", "Question")
    if ($res -eq "No") { return }

    $btnSign.Enabled=$false; $btnAutoSelect.Enabled=$false; $txtSearch.Enabled=$false; $dropZone.Enabled=$false
    $progressBar.Maximum = $totalOps; $progressBar.Value = 0
    $succ=0; $fail=0; $errs=@(); $created=@(); $curr=0

    foreach ($filePath in $script:SelectedFilePaths) {
        # Берем сертификаты ТОЛЬКО для этого файла
        $certs = $script:FileCertMap[$filePath]
        if (-not $certs -or $certs.Count -eq 0) { continue }

        $fileName = Split-Path $filePath -Leaf
        $noExt = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
        $ext = [System.IO.Path]::GetExtension($filePath).TrimStart('.')
        $dir = [System.IO.Path]::GetDirectoryName($filePath)

        foreach ($thumb in $certs) {
            $curr++
            $certObj = $script:AllCertificates | Where-Object { $_.Thumbprint -eq $thumb } | Select-Object -First 1
            $signerName = $certObj.CleanName -replace '[\\/:*?"<>|]', '_' -replace '\s+', ' '
            if ($signerName.Length -gt 40) { $signerName = $signerName.Substring(0,40).Trim() }
            $outFile = Join-Path $dir "${noExt}_${signerName}_${ext}.sig"
            $sgnTmp = $filePath + ".sgn"

            $lblProgress.Text = "[$curr/$totalOps] Файл: $fileName | Подписант: $signerName"
            $lblStatus.Text = "Подписание..."
            $form.Refresh()

            try {
                $p = New-Object System.Diagnostics.Process
                $p.StartInfo.FileName = $CryptCPPath
                $p.StartInfo.Arguments = "-signf -dir `"$dir`" -u -cert -thumbprint `"$thumb`" `"$fileName`""
                $p.StartInfo.RedirectStandardOutput = $true
                $p.StartInfo.RedirectStandardError = $true
                $p.StartInfo.UseShellExecute = $false
                $p.StartInfo.CreateNoWindow = $true
                $p.StartInfo.WorkingDirectory = $dir
                $p.Start()
                $out = $p.StandardOutput.ReadToEnd() + $p.StandardError.ReadToEnd()
                $p.WaitForExit()

                if ($p.ExitCode -eq 0 -and (Test-Path $sgnTmp)) {
                    Start-Sleep -Milliseconds 200
                    $c=1; $base=$outFile
                    while(Test-Path $outFile) { $outFile = "${base}_$c.sig"; $c++ }
                    Rename-Item $sgnTmp $outFile -Force
                    $created += $outFile; $succ++
                } else {
                    $fail++; $errs += "$fileName ($signerName) - Ошибка $($p.ExitCode)`n$out"
                }
            } catch { $fail++; $errs += "$fileName ($signerName) - $_" }
            $progressBar.Value++
        }
    }

    $btnSign.Enabled=$true; $btnAutoSelect.Enabled=$true; $txtSearch.Enabled=$true; $dropZone.Enabled=$true
    $lblStatus.Text = "Готово! Успешно: $succ | Ошибок: $fail"
    $msg = "Подписание завершено!`nУспешно: $succ`nОшибок: $fail"
    if ($created) { $msg += "`n`nСозданные файлы:`n" + ($created -join "`n") }
    if ($errs) { $msg += "`n`n--- ОШИБКИ ---`n" + ($errs -join "`n`n") }
    [System.Windows.Forms.MessageBox]::Show($msg, "Результат", "OK", "Information")
})

LoadCertificates
$form.ShowDialog()