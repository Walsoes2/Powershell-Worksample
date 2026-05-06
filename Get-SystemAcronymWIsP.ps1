
#INFO
# Scriptet körs i en folder med .\GetUniqueSystemAcronym.ps1
# Scriptet kan inte kolla randomgenerade akronymer som kan dyka upp efter att
# Skapat av Henrik Walsöe Vikström

#requires -Version 5.1

#It may broke on PS7.0+ but it is tried on it. 

Write-Host " - INFO - " -ForegroundColor DarkYellow
Write-Host "Scriptet kan inte ändra på randomiserade akronymer som ännu inte skapats. D.v.s det är inte iterativt."
Write-Host "Men det har en check på redan generarade akronymer och ändrar en ej ännu skapad akronym till en unik "
Write-Host "Laddar man dock in listan av akronymer igen som borde vara fullständig nu så genererars unika akronymer för alla systemnamn."
Write-Host "Scriptet är testat med .xlxs-filen 'System.xlsx'"
Write-Host ""
Write-Host "!!!OBS!!! har man inte modulen ImportExcel tar det ett tag att ladda in scriptet/guit"
Write-Host ""
Write-Host "====================================="
Write-Host "  System Acronym Tool starting..."
Write-Host "  Press ANY key to open the GUI"
Write-Host "====================================="
Write-Host ""


[System.Console]::ReadKey($true) | Out-Null


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- GUI ----------------
$form = New-Object Windows.Forms.Form
$form.Text = "Get unique acronyms from system names"
$form.Size = New-Object Drawing.Size(700, 500)
$form.StartPosition = "CenterScreen"

$btnLoad = New-Object Windows.Forms.Button
$btnLoad.Text = "Load File (.txt|.csv|.xlxs)"
$btnLoad.Location = New-Object Drawing.Point(20,20)
$btnLoad.Size = New-Object Drawing.Size(120,40)

$btnRun = New-Object Windows.Forms.Button
$btnRun.Text = "Run"
$btnRun.Location = New-Object Drawing.Point(160,20)
$btnRun.Size = New-Object Drawing.Size(120,30)
$btnRun.Enabled = $false

$btnPreview = New-Object Windows.Forms.Button
$btnPreview.Text = "Preview Input Data"
$btnPreview.Location = New-Object Drawing.Point(300,20)
$btnPreview.Size = New-Object Drawing.Size(120,30)
$btnPreview.Enabled = $false



$btnExport = New-Object Windows.Forms.Button
$btnExport.Text = "Export Excel"
$btnExport.Location = New-Object Drawing.Point(440,20)
$btnExport.Size = New-Object Drawing.Size(120,30)
$btnExport.Enabled = $false

$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Location = New-Object Drawing.Point(20, 450)
$progressBar.Size = New-Object Drawing.Size(640,20)
$progressBar.Minimum = 0

$logBox = New-Object Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.Location = New-Object Drawing.Point(20,70)
$logBox.Size = New-Object Drawing.Size(640,370)


$form.Controls.AddRange(@(
    $btnLoad,
    $btnRun,
    $btnPreview,
    $btnExport,
    $logBox,
    $progressBar
))

# ---------------- STATE ----------------
$script:data = $null
$script:Comparison = @()
  #track used acronyms globally


function Log($msg) {
    $logBox.AppendText("$msg`r`n")
}

function Fix-Headers($rows) {
    if (-not $rows -or $rows.Count -eq 0) { return $rows }

    $first = $rows[0].PSObject.Properties.Name
    $needsFix = $first -contains "" -or $first -match "^Column\d+$"

    if (-not $needsFix) { return $rows }

    $colCount = $first.Count
    $headers = 1..$colCount | ForEach-Object { "Column$_" }

    return $rows | ForEach-Object {
        $i = 0
        $obj = [ordered]@{}
        foreach ($v in $_.PSObject.Properties.Value) {
            $obj[$headers[$i]] = $v
            $i++
        }
        [PSCustomObject]$obj
    }
}

function Get-ColumnScore($name, $type) {
    $n = $name.ToLower()
    $score = 0

    switch ($type) {
        "name" {
            if ($n -match "namn") { $score += 50 }   
            if ($n -match "system")   { $score += 30 }
            if ($n -match "server") { $score += 20 }
        }
        "acro" {
            if ($n -match "acro|akro|aktro") { $score += 50 }
            if ($n -match "short|abbr")      { $score += 30 }
        }
    }

    if ($n -match "^column\d+$") { $score -= 20 }

    return $score
}

function Load-File {

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "CSV, Excel, Text (*.csv;*.xlsx;*.txt)|*.csv;*.xlsx;*.txt"

    if ($dialog.ShowDialog() -ne "OK") { return }

    $file = $dialog.FileName
    $ext  = [System.IO.Path]::GetExtension($file)

    switch ($ext) {
        ".csv"  { $script:data = Import-Csv $file }
        ".xlsx" {
            if (-not (Get-Module -ListAvailable ImportExcel)) {
                Install-Module ImportExcel -Scope CurrentUser -Force
            }
            Import-Module ImportExcel
            $script:data = Import-Excel $file
        }
        ".txt" {
            $script:data = Get-Content $file | ForEach-Object {
                [PSCustomObject]@{ Column1 = $_ }
            }
        }
    }

    $script:data = Fix-Headers $script:data

    if (-not $script:data) {
        Log "No data loaded."
        return
    }

    Log "File loaded: $file"
    Log "Rows: $($script:data.Count)"

    Log ""
Log "Preview (first 2 rows):"

$script:data | Select-Object -First 2 | ForEach-Object {
    Log ($_ | Out-String)
}


    # ---------------- COLUMN DETECTION ----------------
    $columns = $script:data[0].PSObject.Properties.Name

    $scored = $columns | ForEach-Object {
        [PSCustomObject]@{
            Name      = $_
            NameScore = Get-ColumnScore $_ "name"
            AcroScore = Get-ColumnScore $_ "acro"
        }
    }

    $sysCol = ($scored | Sort-Object NameScore -Descending | Select-Object -First 1).Name
    $acroCol = ($scored | Where-Object Name -ne $sysCol |
                Sort-Object AcroScore -Descending | Select-Object -First 1).Name

    # fallback if weak
    if (($scored | Where Name -eq $sysCol).NameScore -lt 10) {

        Log "Auto-detection weak — manual selection required."

        $sel = $columns |
            ForEach-Object { [PSCustomObject]@{ Column = $_ } } |
            Out-GridView -Title "Select Systemnamn + Systemakronym" -PassThru

        if (-not $sel -or $sel.Count -lt 2) { return }

        $sysCol = $sel[0].Column
        $acroCol = $sel[1].Column
    }

    # ---------------- FINAL DATA ----------------
    $script:data = $script:data | ForEach-Object {
        [PSCustomObject]@{
            Systemnamn    = $_.$sysCol
            Systemakronym = $_.$acroCol
        }
    }

    $btnRun.Enabled = $true
    $btnPreview.Enabled = $true
    $btnExport.Enabled = $true

    Log "Columns mapped: $sysCol -> Systemnamn, $acroCol -> Systemakronym"
}

function Invoke-Processing {
    param(
        [Parameter(Mandatory)]
        [string]$Systemnamn,

        [AllowEmptyString()]
        [AllowNull()]
        [string]$Systemakronym
    )

    $felmeddelandeAkronym = ""

    $firstWord = (($Systemnamn -split '\s+')[0] -replace '[^A-Za-z]', '').ToUpper()
$cleanSystemnamn = $Systemnamn -replace '[åäöÅÄÖ]', ''

    $result = ($cleanSystemnamn.ToUpper() -split '\s+' | ForEach-Object {
        $clean = ($_ -replace '[^A-Za-z]', '')
      if ($clean.Length -gt 0) { $clean[0] }
    }) -join ''


    if (-not $result) {
    $result = ""
    }
    # ---- RULES ----
$result = $result.ToString()

	
# "Akronym med fyra ord unik."
    if($result.Length -ge 4 -and $script:Comparison -notcontains $result) { 
	    
        $result = $result.Substring(0,4)
        $felmeddelandeAkronym = "Akronym med fyra eller flera initialer och akronymen finns ej." }

    
#"Akronym ej unik. Sista versalen inkrementerad.    
    elseif ($result.Length -ge 4) {

            $result = $result.substring(0,4)

	while ($script:Comparison -contains $result) {
            $arr = $result.ToCharArray()
            $i = $arr.Length - 1
            $arr[$i] = if ($arr[$i] -eq 'Z') { 'A' } else { [char]([int]$arr[$i] + 1) }
            $result = (-join $arr)
        } 
        
        $felmeddelandeAkronym = "Akronym som bygger på fyra ord. Sista versalen inkrementerad på grund av att den inte är unik. Samma initial finns redan."

    } 
			

    #"Skifte position från första ord plus ev. randomtecken."

    elseif ($result -and $firstword.Length -ge 1) {
            # Start with system prefix
            $candidate = $result
    
            # Add as many chars from firstword as fit (up to 4 total)
            $remaining = 4 - $candidate.Length
            $candidate += $firstword.Substring(1, $firstword.Length-1)
    
            # If still not 4 chars, pad with random letters
            $remaining = 4 - $candidate.Length
            if ($remaining -gt 0) {
                $candidate += -join (1..$remaining | ForEach-Object {
                    [char](Get-Random -Minimum 65 -Maximum 91)
                }) 
            } 
             
            $candidate = $candidate.substring(0,4)

             while ($script:Comparison -contains $candidate) {
                    $arr = $candidate.ToCharArray()
                    $i = $arr.Length - 1
                    $arr[$i] = if ($arr[$i] -eq 'Z') { 'A' } else { [char]([int]$arr[$i] + 1) }
                    $candidate = (-join $arr)
                } 
            
         
    
        $result = $candidate
        $felmeddelandeAkronym = "Skifte position 1+ från första ord med start från bokstav 2. Upp till 4. Finns inte tillräckligt många tecken från första ordet så läggs ett randomtecken till."
    
    }



       	
   
    #"Random genererade akronym."
	
    
        elseif ($result.Length -ge 1 -and $firstword.Length -eq 1) {

            $prefix = $result

            do {
                $candidate = $prefix + -join (1..(4 - $prefix.Length) | ForEach-Object {
                    [char](Get-Random -Minimum 65 -Maximum 91)
                })
            }
            while ($script:Comparison -contains $candidate)  # Kollar att den randomgenerade akronym inte tar en akronym som redan finns i listan (arrayen Comparison)

            $result = $candidate
			$felmeddelandeAkronym = "Random genererad akronym gjort från ett eller flera sammanhängande ord (men inte fyra)"
			
        }

        else {

                do {
                $result = ""

                for ($i = 0; $i -lt 4; $i++) {
                    $result += [char](Get-Random -Minimum 65 -Maximum 91)
                }
				
			 $felmeddelandeAkronym = "Helt randomgenerad akronym eftersom inga bokstäver finns i systemnamnet. T. ex '81' utan bokstäver."	
            }
            while ($script:Comparison -contains $result)
        }

  

      
        
          
   $script:Comparison += $result

    return [PSCustomObject]@{
        Systemnamn = $Systemnamn
        GammalAkronym = $Systemakronym
        NyAkronym    = $result
        FelMedAkronym = $felmeddelandeAkronym
    }
    }

    # store used acronym
	
	
 
  

function Build-Summary {

    $total = $script:results.Count
    $warnings = ($script:results | Where-Object FelMedAkronym -ne "").Count
    $ok = $total - $warnings

    $duplicates = $script:results |
        Group-Object NyAkronym |
        Where-Object Count -gt 1 |
        Select-Object Name, Count

    return [PSCustomObject]@{
        TotalRecords = $total
        OK = $ok
        Warnings = $warnings
        DuplicateGroups = $duplicates.Count
    }
}




# ---------------- BUTTON EVENTS ----------------
$btnLoad.Add_Click({
    Load-File
})

$btnRun.Add_Click({

    if (-not $script:data) {
        Log "No data loaded."
        return
    }

    $script:results = @()
    $script:Comparison = @()


    $progressBar.Value = 0
    $progressBar.Maximum = $script:data.Count

    foreach ($row in $script:data) {

        $result = Invoke-Processing `
            -Systemnamn $row.Systemnamn `
            -Systemakronym $row.Systemakronym

        # ✔ store clean output ONLY
        $script:results += $result

        Log ($result | Out-String)

        $progressBar.Value++
        [System.Windows.Forms.Application]::DoEvents()
    }

    Log "Processing complete."
})


$btnPreview.Add_Click({

    if (-not $script:data) {
        Log "No data to preview."
        return
    }

    if ($script:data.Count -gt 1000) {
        Log "Showing first 1000 rows only."
        $script:data | Select-Object -First 1000 |
            Out-GridView -Title "Preview Input Data"
    }
    else {
        $script:data | Out-GridView -Title "Preview Input Data"
    }
})


$btnExport.Add_Click({

    if (-not $script:results) {
        Log "Nothing to export."
        return
    }

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Log "ImportExcel not found. Installing..."
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -Confirm:$false
    }

    Import-Module ImportExcel

    # Save dialog
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
    $saveDialog.Title = "Save Report"
    $saveDialog.FileName = "SystemAkronymReport.xlsx"

    if ($saveDialog.ShowDialog() -ne "OK") {
        Log "Export cancelled."
        return
    }

    $path = $saveDialog.FileName

    # ---------------- SUMMARY ----------------
    $summary = Build-Summary

    # ---------------- EXPORT ----------------
    $excel = $script:results | Export-Excel -Path $path `
        -WorksheetName "Results" -AutoSize -PassThru -erroraction SilentlyContinue

    # Summary sheet
    $summary | Export-Excel -ExcelPackage $excel -WorksheetName "Summary" -erroraction SilentlyContinue

    Close-ExcelPackage $excel

    Log "Exported report to: $path"

    # auto open
    Start-Process $path
})

# ---------------- SHOW ----------------
$form.ShowDialog()
