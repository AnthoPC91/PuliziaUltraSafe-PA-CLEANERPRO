Write-Host "QUESTO È IL FILE GIUSTO"

# ============================
#   BLOCCO A - FUNZIONI BASE
# ============================

$BasePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogsPath = Join-Path $BasePath "Logs"
$ReportsPath = Join-Path $BasePath "Reports"

if (!(Test-Path $LogsPath)) { New-Item -ItemType Directory -Path $LogsPath | Out-Null }
if (!(Test-Path $ReportsPath)) { New-Item -ItemType Directory -Path $ReportsPath | Out-Null }

function Write-Log {
    param([string]$Message)
    $LogFile = Join-Path $LogsPath ("Log_" + (Get-Date -Format "yyyyMMdd") + ".txt")
    Add-Content -Path $LogFile -Value ("[" + (Get-Date -Format "HH:mm:ss") + "] " + $Message)
}

# ============================
#   PROTEZIONE COMUNE-SAFE
# ============================

function Is-SafePath {
    param([string]$Path)

    if ($Path.StartsWith("\\")) { return $false }
    if ($Path -match '^[A-Z]:\\') {
        $drive = $Path.Substring(0,1)
        if ($drive -ne 'C') { return $false }
    }
    if ($Path -match '\.(mdb|db|sqlite|xml|ini|config)$') { return $false }
    if ($Path -match 'Halley|Maggioli|Sicra|Protocollo|Anagrafe|Tributi|Contabilita') { return $false }

    return $true
}

# ============================
#   FUNZIONI DI PULIZIA (A)
# ============================

function Clean-TempUser {
    Write-Log 'Pulizia TEMP utente'
    Get-ChildItem -Path $env:TEMP -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
        if (Is-SafePath $_.FullName) {
            Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue
        } else {
            Write-Log "SKIPPED (Comune-Safe): $($_.FullName)"
        }
    }
}

function Clean-TempWindows {
    Write-Log 'Pulizia C:\Windows\Temp'
    Get-ChildItem -Path 'C:\Windows\Temp' -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
        if (Is-SafePath $_.FullName) {
            Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue
        } else {
            Write-Log "SKIPPED (Comune-Safe): $($_.FullName)"
        }
    }
}

function Clean-Browsers {
    Write-Log 'Pulizia cache browser'

    $BrowserPaths = @(
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
        "$env:APPDATA\Mozilla\Firefox\Profiles"
    )

    foreach ($path in $BrowserPaths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                if (Is-SafePath $_.FullName) {
                    Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue
                } else {
                    Write-Log "SKIPPED (Comune-Safe): $($_.FullName)"
                }
            }
        }
    }
}
function Clean-WindowsUpdateCache {
    Write-GuiLog "Pulizia cache Windows Update..."
    Write-Log "Pulizia cache Windows Update"

    $path = "C:\Windows\SoftwareDistribution\Download"
    if (Test-Path $path) {
        Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            if (Is-SafePath $_.FullName) {
                Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue
            } else {
                Write-Log "SKIPPED (Comune-Safe): $($_.FullName)"
            }
        }
    }
}

function Clean-OfficeCache {
    Write-GuiLog "Pulizia cache Office..."
    Write-Log "Pulizia cache Office"

    $paths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache",
        "$env:LOCALAPPDATA\Microsoft\Office\15.0\OfficeFileCache"
    )

    foreach ($p in $paths) {
        if (Test-Path $p) {
            Get-ChildItem -Path $p -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                if (Is-SafePath $_.FullName) {
                    Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue
                } else {
                    Write-Log "SKIPPED (Comune-Safe): $($_.FullName)"
                }
            }
        }
    }
}

function Clean-TeamsCache {
    Write-GuiLog "Pulizia cache Teams..."
    Write-Log "Pulizia cache Teams"

    $path = "$env:APPDATA\Microsoft\Teams"
    if (Test-Path $path) {
        Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            if (Is-SafePath $_.FullName) {
                Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue
            } else {
                Write-Log "SKIPPED (Comune-Safe): $($_.FullName)"
            }
        }
    }
}

# ============================
#   BLOCCO B - GUI COMPLETA (COME IERI)
# ============================

Add-Type -AssemblyName PresentationFramework

$XAML = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Title='Pulizia Ultra Safe'
        Height='600' Width='750'
        WindowStartupLocation='CenterScreen'
        Icon='$BasePath\PuliziaUltraSafe.ico'>

    <Grid Margin='20'>

    <!-- STYLE PULSANTI (STEP 5) -->
    <Grid.Resources>

        <Style TargetType="Button">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect Color="Black"
                                      BlurRadius="10"
                                      ShadowDepth="2"
                                      Opacity="0.35"/>
                </Setter.Value>
            </Setter>

            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Opacity" Value="0.85"/>
                    <Setter Property="Effect">
                        <Setter.Value>
                            <DropShadowEffect Color="Black"
                                              BlurRadius="15"
                                              ShadowDepth="3"
                                              Opacity="0.50"/>
                        </Setter.Value>
                    </Setter>
                </Trigger>
            </Style.Triggers>
        </Style>

    </Grid.Resources>

    <!-- Sfondo sfocato -->
    <Image Source='background.jpg'
           Stretch='UniformToFill'>
        <Image.Effect>
            <BlurEffect Radius='18'/>
        </Image.Effect>
    </Image>

    <!-- Pannello chiaro sopra lo sfondo -->
    <Border Background='#CCFFFFFF'
        CornerRadius='10'
        Margin='0,60,0,60'>
    <Border.Effect>
        <DropShadowEffect Color='Black'
                          BlurRadius='20'
                          ShadowDepth='0'
                          Opacity='0.35'/>
    </Border.Effect>
</Border>

    <!-- HEADER ISTITUZIONALE -->
<Border Background='#1A4FB8'
        Height='60'
        VerticalAlignment='Top'
        CornerRadius='10'
        Margin='0,0,0,10'>
    <Border.Effect>
        <DropShadowEffect Color='Black'
                          BlurRadius='15'
                          ShadowDepth='2'
                          Opacity='0.40'/>
    </Border.Effect>
    <Grid>
        <TextBlock Text='Pulizia Ultra Safe-PA-CLEANER PRO'
                   Foreground='White'
                   FontSize='22'
                   FontWeight='Bold'
                   VerticalAlignment='Center'
                   HorizontalAlignment='Center'/>
    </Grid>
</Border>

        <!-- Campo Ufficio -->
        <StackPanel Orientation='Vertical'
                    HorizontalAlignment='Left'
                    VerticalAlignment='Top'
                    Margin='0,120,0,0'>
            <TextBlock Text='Ufficio / Settore' FontWeight='Bold' FontSize='14'/>
            <TextBox x:Name='txtUfficio' Width='300' Height='28' Margin='0,5,0,0'/>
        </StackPanel>

        <!-- Campo Operatore -->
        <StackPanel Orientation='Vertical'
                    HorizontalAlignment='Left'
                    VerticalAlignment='Top'
                    Margin='0,180,0,0'>
            <TextBlock Text='Operatore' FontWeight='Bold' FontSize='14'/>
            <TextBox x:Name='txtOperatore' Width='300' Height='28' Margin='0,5,0,0'/>
        </StackPanel>

        <!-- Checkbox -->
        <CheckBox x:Name='chkAvanzata'
                  Content='Attiva pulizia avanzata sicura'
                  HorizontalAlignment='Left'
                  VerticalAlignment='Top'
                  Margin='0,240,0,0'
                  FontSize='14'/>

        <!-- Log + Progress -->
        <StackPanel HorizontalAlignment='Stretch'
                    VerticalAlignment='Top'
                    Margin='0,290,0,0'>
            <TextBlock Text='Log operazioni' FontWeight='Bold' FontSize='14'/>
            <TextBox x:Name='txtLog'
                     Height='180'
                     Margin='0,5,0,0'
                     TextWrapping='Wrap'
                     VerticalScrollBarVisibility='Auto'
                     AcceptsReturn='True'/>

            <ProgressBar x:Name='progressBar'
                         Minimum='0'
                         Maximum='100'
                         Height='25'
                         Margin='0,10,0,0'
                         Value='0'/>
        </StackPanel>

        <!-- Pulsanti -->
        <StackPanel Orientation='Horizontal'
                    HorizontalAlignment='Center'
                    VerticalAlignment='Bottom'
                    Margin='0,0,0,10'>

            <Button x:Name='btnPulizia'
        Content='AVVIA PULIZIA'
        Width='150' Height='50'
        Margin='10'
        Background='#14AFF8'
        Foreground='White'
        FontWeight='Bold'>
    <Button.Effect>
        <DropShadowEffect Color='Black'
                          BlurRadius='10'
                          ShadowDepth='2'
                          Opacity='0.35'/>
    </Button.Effect>
</Button>


            <Button x:Name='btnAvanzata'
        Content='PULIZIA AVANZATA'
        Width='150' Height='50'
        Margin='10'
        Background='#14AFF8'
        Foreground='White'
        FontWeight='Bold'>
    <Button.Effect>
        <DropShadowEffect Color='Black'
                          BlurRadius='10'
                          ShadowDepth='2'
                          Opacity='0.35'/>
    </Button.Effect>
</Button>

            <Button x:Name='btnReport'
        Content='GENERA REPORT'
        Width='150' Height='50'
        Margin='10'
        Background='#14AFF8'
        Foreground='White'
        FontWeight='Bold'>
    <Button.Effect>
        <DropShadowEffect Color='Black'
                          BlurRadius='10'
                          ShadowDepth='2'
                          Opacity='0.35'/>
    </Button.Effect>
</Button>

            <Button x:Name='btnLog'
        Content='APRI LOG'
        Width='150' Height='50'
        Margin='10'
        Background='#1A4FB8'
        Foreground='White'
        FontWeight='Bold'>
    <Button.Effect>
        <DropShadowEffect Color='Black'
                          BlurRadius='10'
                          ShadowDepth='2'
                          Opacity='0.35'/>
    </Button.Effect>
</Button>

            <Button x:Name='btnPdf'
        Content='ESPORTA PDF'
        Background='#0078D7'
        Foreground='White'
        FontWeight='Bold'
        Padding='10,5'
        Margin='5'
        BorderThickness='0'
        Width='150'
        Height='40'
        Cursor='Hand'>
    <Button.Effect>
        <DropShadowEffect Color='Black'
                          BlurRadius='10'
                          ShadowDepth='2'
                          Opacity='0.35'/>
    </Button.Effect>
</Button>
        </StackPanel>

    </Grid>

</Window>
"@

$bytes = [System.Text.Encoding]::UTF8.GetBytes($XAML)
$stream = New-Object System.IO.MemoryStream
$stream.Write($bytes, 0, $bytes.Length)
$stream.Position = 0

$Window = [Windows.Markup.XamlReader]::Load($stream)
# Imposta sfondo immagine da codice
$imgPath = Join-Path $BasePath 'background.jpg'
$uri = New-Object System.Uri($imgPath)
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage($uri)
$brush = New-Object System.Windows.Media.ImageBrush($bitmap)
$brush.Stretch = 'UniformToFill'
$Window.Background = $brush

$txtUfficio = $Window.FindName('txtUfficio')
$txtOperatore = $Window.FindName('txtOperatore')

$btnPulizia = $Window.FindName('btnPulizia')
$btnAvanzata = $Window.FindName('btnAvanzata')
$btnReport  = $Window.FindName('btnReport')
$btnLog     = $Window.FindName('btnLog')
$btnPdf = $Window.FindName('btnPdf')
$txtLog     = $Window.FindName('txtLog')
$progressBar = $Window.FindName('progressBar')

# ============================
#   BLOCCO C - EVENTI GUI
# ============================

function Write-GuiLog {
    param([string]$msg)
    $txtLog.AppendText("$msg`n")
}
function Set-Progress {
    param([int]$value)
    $progressBar.Value = $value
}
function Export-Pdf {
    param(
        [string]$HtmlContent,
        [string]$OutputPath
    )

    $tempHtml = Join-Path $env:TEMP "temp_export.html"
    Set-Content -Path $tempHtml -Value $HtmlContent -Encoding UTF8

    Write-GuiLog "DEBUG: HTML temporaneo: $tempHtml"

    # NIENTE virgolette dentro l'argomento
    $pdfArg = "--print-to-pdf=$OutputPath"

    # VIRGOLETTE SOLO SULL'HTML
    $quotedHtml = '"' + $tempHtml + '"'

    $args = @(
        "--headless"
        "--disable-gpu"
        "--no-sandbox"
        $pdfArg
        $quotedHtml
    )

    Start-Process -FilePath "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" `
        -ArgumentList $args -NoNewWindow -Wait

    Write-GuiLog "DEBUG: PDF dovrebbe essere qui: $OutputPath"
}
function Export-DeepScanToPdf {
    param(
        [string]$DeepScanContent
    )

    # Percorso PDF (puoi cambiarlo se vuoi)
    $pdfPath = Join-Path $env:TEMP ("DeepScan_" + (Get-Date -Format 'yyyyMMdd_HHmm') + ".pdf")

    # HTML semplice
    $html = @"
<html>
<head>
<meta charset='UTF-8'>
<style>
body { font-family: Arial; font-size: 12pt; }
h1 { color: #003366; }
pre { background: #f0f0f0; padding: 10px; border-radius: 5px; }
</style>
</head>
<body>
<h1>Report Scansione Profonda</h1>
<pre>
$DeepScanContent
</pre>
</body>
</html>
"@

    Export-Pdf -HtmlContent $html -OutputPath $pdfPath

    Write-GuiLog "PDF Scansione Profonda creato: $pdfPath"
return $pdfPath
}
function Export-LogToPdf {
    $logText = $txtLog.Text -replace "`n", "<br>"

    $html = @"
<html>
<body style='font-family:Arial'>
<h2>Log Operazioni</h2>
<p>$logText</p>
</body>
</html>
"@

    $pdfPath = Join-Path $env:TEMP ("Log_$(Get-Date -Format 'yyyyMMdd_HHmm').pdf")
    Export-Pdf -HtmlContent $html -OutputPath $pdfPath

    Write-GuiLog "PDF log esportato: $pdfPath"
}
function Export-ReportToPdf {
    $latestReport = Get-ChildItem -Path $ReportsPath -Filter *.html | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if (-not $latestReport) {
        Write-GuiLog "Nessun report trovato"
        return
    }

    $html = Get-Content $latestReport.FullName -Raw
    $pdfPath = Join-Path $env:TEMP ("Log_$(Get-Date -Format 'yyyyMMdd_HHmm').pdf")

    Export-Pdf -HtmlContent $html -OutputPath $pdfPath

    Write-GuiLog "PDF report esportato: $pdfPath"
}

$btnPulizia.Add_Click({
    Set-Progress 0
    Write-GuiLog "Pulizia avviata"

    Set-Progress 20
    Write-GuiLog "Pulizia TEMP utente..."
    Clean-TempUser

    Set-Progress 50
    Write-GuiLog "Pulizia Windows Temp..."
    Clean-TempWindows

    Set-Progress 80
    Write-GuiLog "Pulizia cache browser..."
    Clean-Browsers

    Set-Progress 100
    Write-GuiLog "Pulizia completata"
})

$btnAvanzata.Add_Click({
    Set-Progress 0
    Write-GuiLog "Pulizia avanzata avviata"
    $deepScanResult = "Pulizia avanzata avviata"

    # ============================
    #   1) PULIZIA NORMALE
    # ============================

    Set-Progress 10
    Write-GuiLog "Pulizia TEMP utente..."
    Clean-TempUser
    $deepScanResult += "`nPulizia TEMP utente completata"

    Set-Progress 25
    Write-GuiLog "Pulizia Windows Temp..."
    Clean-TempWindows
    $deepScanResult += "`nPulizia Windows Temp completata"

    Set-Progress 40
    Write-GuiLog "Pulizia cache browser..."
    Clean-Browsers
    $deepScanResult += "`nPulizia cache browser completata"

    # ============================
    #   2) PULIZIA AVANZATA
    # ============================

    Set-Progress 55
    Write-GuiLog "Pulizia cache Windows Update..."
    Clean-WindowsUpdateCache
    $deepScanResult += "`nPulizia cache Windows Update completata"

    Set-Progress 70
    Write-GuiLog "Pulizia cache Office..."
    Clean-OfficeCache
    $deepScanResult += "`nPulizia cache Office completata"

    Set-Progress 85
    Write-GuiLog "Pulizia cache Teams..."
    Clean-TeamsCache
    $deepScanResult += "`nPulizia cache Teams completata"

    # ============================
    #   3) FINALE + PDF
    # ============================

    Set-Progress 100
    Write-GuiLog "Pulizia avanzata completata"
    $deepScanResult += "`nPulizia avanzata completata"

    # CREA IL PDF E RICEVE IL PERCORSO
    $pdfPath = Export-DeepScanToPdf -DeepScanContent $deepScanResult

    # APRE DIRETTAMENTE IL PDF
    Start-Process $pdfPath
})

$btnReport.Add_Click({
    $ufficio = $txtUfficio.Text
    $operatore = $txtOperatore.Text
    $data = Get-Date -Format 'dd/MM/yyyy HH:mm'

    # Tipo di pulizia (normale o avanzata)
    if ($chkAvanzata.IsChecked) {
        $tipo = "Pulizia Avanzata Sicura"
    } else {
        $tipo = "Pulizia Standard"
    }

    $ReportFile = Join-Path $ReportsPath ("Report_" + (Get-Date -Format 'yyyyMMdd_HHmm') + '.html')

    $html = @"
<html>
<body style='font-family:Arial'>
<h2>Report Pulizia Ultra Safe</h2>

<p><b>Ufficio / Settore:</b> $ufficio</p>
<p><b>Operatore:</b> $operatore</p>
<p><b>Data:</b> $data</p>
<p><b>Tipo di pulizia:</b> $tipo</p>

<hr>

<p>La pulizia è stata eseguita correttamente.</p>

</body>
</html>
"@

    Set-Content -Path $ReportFile -Value $html -Encoding UTF8

    Write-Log "Report generato: $ReportFile"
    Write-GuiLog "Report generato"
})

$btnLog.Add_Click({
    Invoke-Item $LogsPath
})
$btnPdf.Add_Click({
    Export-LogToPdf
    Export-ReportToPdf
})

$Window.ShowDialog() | Out-Null
