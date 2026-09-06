# ColaImpresion ECP - V13 PDF PORTABLE - CORREGIDO
# Base: version que convierte DOC/DOCX a PDF antes de imprimir (evita fallos de COM/Word en equipos ajenos).
# FIX aplicado: $nombreImpresora no llegaba a Imprimir-Word (estaba fuera de alcance), por lo que
# TODO documento Word fallaba al imprimir sin importar el equipo. Ahora se pasa como parametro explicito.
#
# IMPORTANTE - Estructura de carpetas requerida (distribuir COMPLETA, nunca el .exe suelto):
#   ColaImpresion
#       ColaImpresion.exe
#       LogoECP.png
#       perro2.png
#       MotorPDF
#           PDFtoImage.dll
#           SkiaSharp.dll
#           System.Memory.dll
#           System.Buffers.dll
#           System.Runtime.CompilerServices.Unsafe.dll
#           pdfium.dll
#           libSkiaSharp.dll
#
# IMPORTANTE - Compilacion con PS2EXE:
#   ps2exe -inputFile ColaImpresion_V13_PORTABLE_CORREGIDO.ps1 -outputFile ColaImpresion.exe -x64 -noConsole
#   Verifica que pdfium.dll y libSkiaSharp.dll sean tambien build x64 (mismatch de arquitectura
#   es la causa mas comun de que el motor PDF no cargue en equipos distintos al tuyo).
#
# IMPORTANTE - Requiere Microsoft Word instalado en el equipo (se usa para convertir DOC/DOCX a PDF
# via VBScript). Si el equipo destino no tiene Word, la impresion de Word fallara (no es bug, es
# requisito de esta funcionalidad).

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
# MOTOR PDF INTEGRADO - PDFtoImage / PDFium
# ============================================================

$script:PdfEngineReady = $false
$script:PdfEngineMessage = ""
$script:PdfDiagnostic = ""

function Inicializar-Motor-PDF {

    # Primero intentamos la carpeta del script. Si estamos dentro de un EXE
    # generado con PS2EXE, PSScriptRoot puede apuntar a una carpeta temporal;
    # en ese caso usamos la carpeta real donde esta el EXE.
    $basePathECP = $PSScriptRoot

    if ([string]::IsNullOrWhiteSpace($basePathECP) -or
        -not (Test-Path (Join-Path $basePathECP "MotorPDF"))) {
        try {
            $ejecutable = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
            $directorioExe = Split-Path -Parent $ejecutable
            if (Test-Path (Join-Path $directorioExe "MotorPDF")) {
                $basePathECP = $directorioExe
            }
        } catch {}
    }

    if ([string]::IsNullOrWhiteSpace($basePathECP) -or
        -not (Test-Path (Join-Path $basePathECP "MotorPDF"))) {
        $basePathECP = [System.Windows.Forms.Application]::StartupPath
    }

    if ([string]::IsNullOrWhiteSpace($basePathECP)) {
        $basePathECP = (Get-Location).Path
    }

    $carpetaArquitectura = if ([System.Environment]::Is64BitProcess) { "x64" } else { "x86" }
    $pdfEngineDir = Join-Path (Join-Path $basePathECP "MotorPDF") $carpetaArquitectura

    if (!(Test-Path $pdfEngineDir)) {
        $pdfEngineDirPlano = Join-Path $basePathECP "MotorPDF"
        if (Test-Path $pdfEngineDirPlano) {
            $pdfEngineDir = $pdfEngineDirPlano
        }
    }

    if (!(Test-Path $pdfEngineDir)) {
        $script:PdfEngineMessage = "No se encontro el motor PDF en: $pdfEngineDir`r`n`r`nAsegurate de distribuir la carpeta MotorPDF\$carpetaArquitectura junto al ejecutable (proceso detectado como $carpetaArquitectura)."
        return $false
    }

    try {

        # Desbloqueo defensivo: si la carpeta MotorPDF se copio desde un origen
        # marcado como "Internet" (descarga, zip, Teams, correo, etc.), Windows
        # puede haber marcado TODAS las DLL (incluidas las nativas pdfium.dll y
        # libSkiaSharp.dll) con "Mark of the Web". Assembly.Load(bytes) evita el
        # bloqueo para las DLL administradas, pero las nativas se cargan por el
        # loader de Windows (via PATH), asi que las desbloqueamos aqui tambien
        # para prevenir problemas silenciosos.
        try {
            Get-ChildItem -Path $pdfEngineDir -Filter "*.dll" -ErrorAction SilentlyContinue |
                Unblock-File -ErrorAction SilentlyContinue
        } catch {}

        $env:PATH = $pdfEngineDir + ";" + $env:PATH

        $managedNames = @(
            "System.Memory.dll",
            "System.Buffers.dll",
            "System.Runtime.CompilerServices.Unsafe.dll",
            "SkiaSharp.dll",
            "PDFtoImage.dll"
        )

        foreach ($name in $managedNames) {
            $dllPath = Join-Path $pdfEngineDir $name

            if (!(Test-Path $dllPath)) {
                throw "No se encontro la dependencia requerida: $dllPath"
            }

            $assemblyName = [System.IO.Path]::GetFileNameWithoutExtension($name)
            $yaCargada = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq $assemblyName } | Select-Object -First 1

            if ($null -eq $yaCargada) {
                # Se usa Load(bytes) en vez de LoadFrom(ruta) para evitar el bloqueo de
                # "Mark of the Web" (HRESULT 0x80131515 / COR_E_LOADFROMBLOCKED). Cuando
                # la carpeta MotorPDF se copia desde un origen marcado como "Internet"
                # (descarga, zip bajado, compartido por Teams/correo, etc.), Windows
                # agrega esa marca a cada DLL y .NET bloquea LoadFrom aunque el archivo
                # sea valido. Cargar desde bytes en memoria evita esa verificacion.
                $ensamblado = [System.Reflection.Assembly]::Load([System.IO.File]::ReadAllBytes($dllPath))
                if ($null -eq $ensamblado) {
                    throw "No se pudo cargar: $dllPath"
                }
            }
        }

        $skiaCargada = [AppDomain]::CurrentDomain.GetAssemblies() |
            Where-Object { $_.GetName().Name -eq "SkiaSharp" } |
            Select-Object -First 1

        if (($null -eq $skiaCargada) -or ($skiaCargada.GetName().Version.ToString() -ne "4.150.0.0")) {
            throw "Version incorrecta de SkiaSharp. PDFtoImage 5.4.0 requiere 4.150.0.0."
        }

        [void][PDFtoImage.Conversion]
        [void][PDFtoImage.RenderOptions]

        $script:PdfEngineReady = $true
        $script:PdfEngineMessage = ""
        return $true
    }
    catch {
        $script:PdfEngineReady = $false
        $mensajeArquitectura = ""
        if ($_.Exception -is [System.BadImageFormatException] -or
            ($null -ne $_.Exception.InnerException -and $_.Exception.InnerException -is [System.BadImageFormatException])) {
            $mensajeArquitectura = "`r`n`r`nPOSIBLE CAUSA: mezcla de arquitecturas x86/x64. Verifica que el .exe se compilo con -x64 y que pdfium.dll / libSkiaSharp.dll tambien sean x64."
        }
        $script:PdfEngineMessage = "Tipo: $($_.Exception.GetType().FullName)`r`nMensaje: $($_.Exception.Message)$mensajeArquitectura"
        if ($null -ne $_.Exception.InnerException) {
            $script:PdfEngineMessage += "`r`nInterno: $($_.Exception.InnerException.Message)"
        }
        return $false
    }
}

[void](Inicializar-Motor-PDF)


# ============================================================
# CONFIGURACION REMOTA FIRMADA - ECP
# ============================================================
# La CLABE, textos de apoyo, interruptor de activacion y version
# minima requerida se descargan de una configuracion oficial HTTPS
# y se verifican con una firma digital RSA antes de confiar en ellos.
#
# Ver Generar-Llaves-ECP.ps1 y Firmar-Config-ECP.ps1 para el flujo
# completo de generacion de llaves y firmado del JSON.
# ============================================================

# Version de ESTA build del .exe/.ps1. Actualiza este numero cada vez
# que compiles una nueva version, y sube tambien el campo "version"
# del JSON remoto para que coincidan.
$script:VersionAppECP = "1.2.0"

$script:UrlConfigECP = "https://wishlly.web.app/soporte/cola-impresion.json"

# ------------------------------------------------------------
# LLAVE PUBLICA - segura de compartir/embeber, solo verifica firmas.
# PENDIENTE: reemplaza el valor de ejemplo de abajo por el contenido
# real de tu archivo ECP_LlavePublica.xml (generado con
# Generar-Llaves-ECP.ps1). Debe quedar como una sola linea XML
# tipo <RSAKeyValue><Modulus>...</Modulus><Exponent>...</Exponent></RSAKeyValue>
# ------------------------------------------------------------
$script:LlavePublicaECP = @"
<RSAKeyValue><Modulus>3V7r0Je1Hm4Z0oJ7+b6pmtVFtq5lFInzUC2KsB7oNT8R32ywQasxuNiP38+kg3NVh+MS/9wIUHN/asY2v64DiFJGMeFYzYmXNCXCmiltGqr2QKx3o2n9qnCKjUDpY/SM6dkl2eww9il4cR9mPc+UIZlf7itnOFjMIhhRZb1QXM09F/bZ54JE7UZmSZHlcBCHfk8qdL9GFrQvLjP8WKj+PrD2BGZaNZKheLP68H71zeRz7NOIr3NvnmOkzIddsu5YiuEpueyCS05V6cEg3D5xaqvhGMsDe4RftFryAEtI1eLdVpbDNGepwGuWZ1zyYNpK/wCJ6xsgpI0FmL+2/bVqqQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>
"@

# ------------------------------------------------------------
# RESPALDO LOCAL - se usa unicamente si no hay internet/servidor
# no disponible, o si la firma del JSON remoto no es valida.
# ------------------------------------------------------------
$script:ConfigRespaldoECP = [ordered]@{
    version              = $script:VersionAppECP
    minVersionRequerida  = "1.0.0"
    activo               = $true
    titulo               = "¿Te ahorró unos minutos de sufrimiento?"
    mensaje              = "esta herramienta es gratis para todos los godínez."
    aviso                = "No es obligatoria, tu jefe ya te pide suficiente"
    beneficiario         = "Edgar Antonio Chirinos Pérez"
    banco                = "Klar"
    clabe                = "661180006850132672"
}

# ------------------------------------------------------------
# ORDEN CANONICO DE CAMPOS PARA LA FIRMA - debe coincidir EXACTO
# con el mismo arreglo en Firmar-Config-ECP.ps1. No cambiar sin
# actualizar tambien el firmador.
# ------------------------------------------------------------
$script:CamposCanonicosECP = @(
    "version","minVersionRequerida","activo","titulo","mensaje","aviso","beneficiario","banco","clabe"
)
$script:SeparadorCanonicoECP = [char]1

function Construir-CadenaCanonica-ECP {
    param([object]$config)
    $partes = New-Object System.Collections.Generic.List[string]
    foreach ($campo in $script:CamposCanonicosECP) {
        if (-not ($config.PSObject.Properties.Name -contains $campo)) {
            throw "Falta el campo requerido en la configuracion: '$campo'"
        }
        $valor = $config.$campo
        if ($campo -eq "activo") {
            $valorTexto = ([bool]$valor).ToString().ToLowerInvariant()
        } else {
            $valorTexto = [string]$valor
        }
        [void]$partes.Add($valorTexto)
    }
    return ($partes -join $script:SeparadorCanonicoECP)
}

function Verificar-Firma-ECP {
    param([object]$config)

    if (-not ($config.PSObject.Properties.Name -contains "firma") -or
        [string]::IsNullOrWhiteSpace($config.firma)) {
        return $false
    }

    try {
        $cadenaCanonica = Construir-CadenaCanonica-ECP $config
        $bytesCanonicos = [System.Text.Encoding]::UTF8.GetBytes($cadenaCanonica)
        $firmaBytes = [Convert]::FromBase64String($config.firma)

        $cspParams = New-Object System.Security.Cryptography.CspParameters
        $cspParams.ProviderType = 24  # PROV_RSA_AES, soporta SHA-256

        $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider(2048, $cspParams)
        try {
            $rsa.FromXmlString($script:LlavePublicaECP)
            return $rsa.VerifyData($bytesCanonicos, "SHA256", $firmaBytes)
        } finally {
            $rsa.Dispose()
        }
    } catch {
        return $false
    }
}

function Version-EsMenorQue-ECP {
    param([string]$versionActual, [string]$versionMinima)
    try {
        return ([version]$versionActual) -lt ([version]$versionMinima)
    } catch {
        # Si algun valor no se puede interpretar como version, no bloqueamos por esto.
        return $false
    }
}

function Descargar-Configuracion-ECP {
    try {
        $cliente = New-Object System.Net.WebClient
        $cliente.Headers["User-Agent"] = "ColaImpresion-ECP/$($script:VersionAppECP)"
        $jsonBytes = $cliente.DownloadData($script:UrlConfigECP)
        $jsonTexto = [System.Text.Encoding]::UTF8.GetString($jsonBytes)
        if ([string]::IsNullOrWhiteSpace($jsonTexto)) { throw "Respuesta vacia del servidor." }
        $config = $jsonTexto | ConvertFrom-Json
        return @{ Exito = $true; Config = $config }
    } catch {
        return @{ Exito = $false; Config = $null; Error = $_.Exception.Message }
    }
}

# $script:ConfigAdvertenciaECP puede ser: "" (sin problema), "FirmaInvalida"
# (posible manipulacion) o "VersionDesactualizada".
$script:ConfigAdvertenciaECP = ""

function Resolver-Configuracion-ECP {

    $descarga = Descargar-Configuracion-ECP

    if (-not $descarga.Exito) {
        # Sin internet / servidor no disponible: esto NO es manipulacion,
        # asi que no se muestra ninguna advertencia y la app sigue normal
        # usando el respaldo local.
        return [pscustomobject]$script:ConfigRespaldoECP
    }

    $config = $descarga.Config
    $firmaValida = Verificar-Firma-ECP $config

    if (-not $firmaValida) {
        # El JSON se descargo pero su firma no coincide con la llave publica:
        # esto SI es una señal de posible manipulacion (hosting comprometido,
        # ataque en transito, o edicion manual sin la llave privada).
        # Nunca se usan estos datos (ni la CLABE ni los textos); se cae al
        # respaldo local de confianza.
        $script:ConfigAdvertenciaECP = "FirmaInvalida"
        return [pscustomobject]$script:ConfigRespaldoECP
    }

    if (Version-EsMenorQue-ECP -versionActual $script:VersionAppECP -versionMinima ([string]$config.minVersionRequerida)) {
        $script:ConfigAdvertenciaECP = "VersionDesactualizada"
    }

    return $config
}

function Mostrar-Advertencia-Seguridad-ECP {
    param([string]$mensaje)

    $formAdvertencia = New-Object System.Windows.Forms.Form
    $formAdvertencia.Text = "Advertencia de seguridad - ECP"
    $formAdvertencia.Size = New-Object System.Drawing.Size(540,340)
    $formAdvertencia.StartPosition = "CenterScreen"
    $formAdvertencia.BackColor = [System.Drawing.Color]::FromArgb(18,18,18)
    $formAdvertencia.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)
    $formAdvertencia.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $formAdvertencia.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $formAdvertencia.MaximizeBox = $false
    $formAdvertencia.MinimizeBox = $false
    $formAdvertencia.TopMost = $true

    $lblIconoAdv = New-Object System.Windows.Forms.Label
    $lblIconoAdv.Text = "⚠"
    $lblIconoAdv.Font = New-Object System.Drawing.Font("Segoe UI",30,[System.Drawing.FontStyle]::Bold)
    $lblIconoAdv.ForeColor = [System.Drawing.Color]::FromArgb(240,170,60)
    $lblIconoAdv.AutoSize = $true
    $lblIconoAdv.Location = New-Object System.Drawing.Point(25,20)
    $formAdvertencia.Controls.Add($lblIconoAdv)

    $lblMensajeAdv = New-Object System.Windows.Forms.Label
    $lblMensajeAdv.Text = $mensaje
    $lblMensajeAdv.Location = New-Object System.Drawing.Point(25,85)
    $lblMensajeAdv.Size = New-Object System.Drawing.Size(485,170)
    $lblMensajeAdv.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)
    $formAdvertencia.Controls.Add($lblMensajeAdv)

    $btnContinuarAdv = New-Object System.Windows.Forms.Button
    $btnContinuarAdv.Text = "Continuar bajo mi riesgo"
    $btnContinuarAdv.Size = New-Object System.Drawing.Size(230,40)
    $btnContinuarAdv.Location = New-Object System.Drawing.Point(25,260)
    $btnContinuarAdv.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
    $btnContinuarAdv.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)
    $btnContinuarAdv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnContinuarAdv.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(65,65,65)
    $btnContinuarAdv.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $formAdvertencia.Controls.Add($btnContinuarAdv)

    $btnCerrarAdv = New-Object System.Windows.Forms.Button
    $btnCerrarAdv.Text = "Cerrar aplicación"
    $btnCerrarAdv.Size = New-Object System.Drawing.Size(230,40)
    $btnCerrarAdv.Location = New-Object System.Drawing.Point(280,260)
    $btnCerrarAdv.BackColor = [System.Drawing.Color]::FromArgb(235,80,80)
    $btnCerrarAdv.ForeColor = [System.Drawing.Color]::White
    $btnCerrarAdv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCerrarAdv.DialogResult = [System.Windows.Forms.DialogResult]::No
    $formAdvertencia.Controls.Add($btnCerrarAdv)

    # Por seguridad, la tecla Enter cierra la aplicacion por defecto,
    # no la continua.
    $formAdvertencia.AcceptButton = $btnCerrarAdv

    $resultadoAdv = $formAdvertencia.ShowDialog()
    $formAdvertencia.Dispose()

    return ($resultadoAdv -eq [System.Windows.Forms.DialogResult]::Yes)
}

$script:ConfigECP = Resolver-Configuracion-ECP

if ($script:ConfigAdvertenciaECP -eq "FirmaInvalida") {
    $seguirAdelante = Mostrar-Advertencia-Seguridad-ECP (
        "No se pudo verificar la autenticidad de la configuración remota de esta aplicación.`r`n`r`n" +
        "Esto puede indicar un problema en tu red, en el servidor, o que el archivo de " +
        "configuración fue modificado sin autorización.`r`n`r`n" +
        "Por tu seguridad, los datos de apoyo/donativo se reemplazaron por un respaldo local " +
        "de confianza. La función de impresión no se ve afectada."
    )
    if (-not $seguirAdelante) { return }
}
elseif ($script:ConfigAdvertenciaECP -eq "VersionDesactualizada") {
    $seguirAdelante = Mostrar-Advertencia-Seguridad-ECP (
        "Esta versión de Cola de Impresión ECP está desactualizada (versión actual: $($script:VersionAppECP)).`r`n`r`n" +
        "Se recomienda actualizar a la versión más reciente antes de continuar usándola."
    )
    if (-not $seguirAdelante) { return }
}


$basePath = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($basePath) -or
    -not (Test-Path (Join-Path $basePath "LogoECP.png"))) {
    try {
        $ejecutable = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $directorioExe = Split-Path -Parent $ejecutable
        if (Test-Path (Join-Path $directorioExe "LogoECP.png")) {
            $basePath = $directorioExe
        }
    } catch {}
}
if ([string]::IsNullOrWhiteSpace($basePath)) {
    $basePath = [System.Windows.Forms.Application]::StartupPath
}
$rutaLogo = Join-Path $basePath "LogoECP.png"
$extensionesPermitidas = @(".pdf",".doc",".docx",".tif",".tiff")

$colorFondo=[System.Drawing.Color]::FromArgb(18,18,18)
$colorPanel=[System.Drawing.Color]::FromArgb(25,25,25)
$colorControl=[System.Drawing.Color]::FromArgb(30,30,30)
$colorBorde=[System.Drawing.Color]::FromArgb(65,65,65)
$colorTexto=[System.Drawing.Color]::FromArgb(235,235,235)
$colorTextoSecundario=[System.Drawing.Color]::FromArgb(185,185,185)
$colorAzul=[System.Drawing.Color]::FromArgb(45,140,255)
$colorVerde=[System.Drawing.Color]::FromArgb(100,210,100)
$colorRojo=[System.Drawing.Color]::FromArgb(235,80,80)
$colorNaranja=[System.Drawing.Color]::FromArgb(240,170,60)

$form=New-Object System.Windows.Forms.Form
$form.Text="Cola de Impresion - ECP"
$form.Size=New-Object System.Drawing.Size(1050,700)
$form.MinimumSize=New-Object System.Drawing.Size(900,600)
$form.StartPosition="CenterScreen"
$form.BackColor=$colorFondo
$form.ForeColor=$colorTexto
$form.Font=New-Object System.Drawing.Font("Segoe UI",10)

$panelLogo=New-Object System.Windows.Forms.Panel
$panelLogo.Location=New-Object System.Drawing.Point(25,15)
$panelLogo.Size=New-Object System.Drawing.Size(105,65)
$panelLogo.BackColor=$colorFondo
$form.Controls.Add($panelLogo)

$radioLogo=12
$rectLogo=New-Object System.Drawing.Rectangle(0,0,$panelLogo.Width,$panelLogo.Height)
$pathLogo=New-Object System.Drawing.Drawing2D.GraphicsPath
$d=$radioLogo*2
$pathLogo.AddArc($rectLogo.X,$rectLogo.Y,$d,$d,180,90)
$pathLogo.AddArc($rectLogo.Right-$d,$rectLogo.Y,$d,$d,270,90)
$pathLogo.AddArc($rectLogo.Right-$d,$rectLogo.Bottom-$d,$d,$d,0,90)
$pathLogo.AddArc($rectLogo.X,$rectLogo.Bottom-$d,$d,$d,90,90)
$pathLogo.CloseFigure()
$panelLogo.Region=New-Object System.Drawing.Region($pathLogo)

$pictureLogo=New-Object System.Windows.Forms.PictureBox
$pictureLogo.Location=New-Object System.Drawing.Point(5,5)
$pictureLogo.Size=New-Object System.Drawing.Size(95,55)
$pictureLogo.SizeMode=[System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureLogo.BackColor=[System.Drawing.Color]::Transparent
if(Test-Path $rutaLogo){try{$pictureLogo.Image=[System.Drawing.Image]::FromFile($rutaLogo)}catch{}}
$panelLogo.Controls.Add($pictureLogo)

$labelTitulo=New-Object System.Windows.Forms.Label
$labelTitulo.Text="COLA DE IMPRESION"
$labelTitulo.Font=New-Object System.Drawing.Font("Segoe UI",18,[System.Drawing.FontStyle]::Bold)
$labelTitulo.ForeColor=$colorTexto
$labelTitulo.AutoSize=$true
$labelTitulo.Location=New-Object System.Drawing.Point(155,32)
$form.Controls.Add($labelTitulo)

$labelImpresora=New-Object System.Windows.Forms.Label
$labelImpresora.Text="Impresora:"
$labelImpresora.AutoSize=$true
$labelImpresora.Location=New-Object System.Drawing.Point(25,95)
$form.Controls.Add($labelImpresora)

$comboImpresoras=New-Object System.Windows.Forms.ComboBox
$comboImpresoras.Location=New-Object System.Drawing.Point(100,91)
$comboImpresoras.Size=New-Object System.Drawing.Size(500,30)
$comboImpresoras.DropDownStyle="DropDownList"
$comboImpresoras.BackColor=$colorControl
$comboImpresoras.ForeColor=$colorTexto
$form.Controls.Add($comboImpresoras)

try{
    $impresoras=Get-CimInstance Win32_Printer|Sort-Object Name
    foreach($impresora in $impresoras){[void]$comboImpresoras.Items.Add($impresora.Name)}
    $predeterminada=$impresoras|Where-Object{$_.Default -eq $true}|Select-Object -First 1
    if($null -ne $predeterminada){$comboImpresoras.SelectedItem=$predeterminada.Name}
    elseif($comboImpresoras.Items.Count -gt 0){$comboImpresoras.SelectedIndex=0}
}catch{
    [System.Windows.Forms.MessageBox]::Show("No fue posible detectar las impresoras instaladas.","Error",
    [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
}

$labelEstado=New-Object System.Windows.Forms.Label
$labelEstado.Text="Estado: Lista para imprimir"
$labelEstado.ForeColor=$colorVerde
$labelEstado.AutoSize=$true
$labelEstado.Location=New-Object System.Drawing.Point(620,95)
$form.Controls.Add($labelEstado)

# ============================================================
# AYUDA DE RANGOS
# ============================================================

$panelAyuda=New-Object System.Windows.Forms.Panel
$panelAyuda.Location=New-Object System.Drawing.Point(25,505)
$panelAyuda.Size=New-Object System.Drawing.Size(785,58)
$panelAyuda.BackColor=[System.Drawing.Color]::FromArgb(24,24,24)
$panelAyuda.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($panelAyuda)

$labelAyuda=New-Object System.Windows.Forms.Label
$labelAyuda.Text="Rangos: vacio = todas  |  2 = solo pag. 2  |  1-3 = pags. 1 a 3  |  1,3 = pag. 1 y pag. 3 por separado  |  1-5,6-6 = bloques independientes"
$labelAyuda.ForeColor=$colorTextoSecundario
$labelAyuda.AutoSize=$false
$labelAyuda.Dock=[System.Windows.Forms.DockStyle]::Fill
$labelAyuda.Padding=New-Object System.Windows.Forms.Padding(10,7,10,5)
$labelAyuda.TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft
$panelAyuda.Controls.Add($labelAyuda)

$labelInfo=New-Object System.Windows.Forms.Label
$labelInfo.Text="Documentos en cola: 0"
$labelInfo.ForeColor=$colorTextoSecundario
$labelInfo.AutoSize=$true
$labelInfo.Location=New-Object System.Drawing.Point(25,130)
$form.Controls.Add($labelInfo)

$tabla=New-Object System.Windows.Forms.DataGridView
$tabla.Location=New-Object System.Drawing.Point(25,160)
$tabla.Size=New-Object System.Drawing.Size(975,335)
$tabla.BackgroundColor=$colorPanel
$tabla.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle
$tabla.GridColor=$colorBorde
$tabla.AllowUserToAddRows=$false
$tabla.AllowUserToDeleteRows=$false
$tabla.AllowUserToResizeRows=$false
$tabla.RowHeadersVisible=$false
$tabla.SelectionMode=[System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$tabla.MultiSelect=$false
$tabla.AutoGenerateColumns=$false
$tabla.EditMode=[System.Windows.Forms.DataGridViewEditMode]::EditOnEnter
$tabla.EnableHeadersVisualStyles=$false
$tabla.ColumnHeadersDefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(35,35,35)
$tabla.ColumnHeadersDefaultCellStyle.ForeColor=$colorTexto
$tabla.ColumnHeadersDefaultCellStyle.Font=New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$tabla.ColumnHeadersDefaultCellStyle.Alignment=[System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$tabla.ColumnHeadersHeight=35
$tabla.DefaultCellStyle.BackColor=$colorPanel
$tabla.DefaultCellStyle.ForeColor=$colorTexto
$tabla.DefaultCellStyle.SelectionBackColor=[System.Drawing.Color]::FromArgb(35,100,175)
$tabla.DefaultCellStyle.SelectionForeColor=[System.Drawing.Color]::White
$tabla.DefaultCellStyle.Font=New-Object System.Drawing.Font("Segoe UI",10)
$tabla.RowTemplate.Height=30
$form.Controls.Add($tabla)

$colNo=New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNo.Name="No";$colNo.HeaderText="No.";$colNo.Width=55;$colNo.ReadOnly=$true
$colNo.SortMode=[System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$colNo.DefaultCellStyle.Alignment=[System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
[void]$tabla.Columns.Add($colNo)

$colArchivo=New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colArchivo.Name="Archivo";$colArchivo.HeaderText="Archivo";$colArchivo.Width=500;$colArchivo.ReadOnly=$true
$colArchivo.SortMode=[System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
[void]$tabla.Columns.Add($colArchivo)

$colPaginas=New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colPaginas.Name="Paginas";$colPaginas.HeaderText="Paginas";$colPaginas.Width=90;$colPaginas.ReadOnly=$true
$colPaginas.SortMode=[System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$colPaginas.DefaultCellStyle.Alignment=[System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
[void]$tabla.Columns.Add($colPaginas)

$colRango=New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colRango.Name="Rango";$colRango.HeaderText="Imprimir de";$colRango.Width=280;$colRango.ReadOnly=$false
$colRango.SortMode=[System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$colRango.DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(32,32,32)
$colRango.DefaultCellStyle.ForeColor=$colorTexto
[void]$tabla.Columns.Add($colRango)

$archivos=New-Object System.Collections.ArrayList

function Procesar-Eventos{[System.Windows.Forms.Application]::DoEvents()}

function Obtener-Paginas-TIFF{
 param([string]$ruta)
 $imagen=$null
 try{
  $imagen=[System.Drawing.Image]::FromFile($ruta)
  $dimension=[System.Drawing.Imaging.FrameDimension]::Page
  $cantidad=$imagen.GetFrameCount($dimension)
  if($cantidad -lt 1){$cantidad=1}
  return $cantidad
 }catch{return 0}
 finally{if($null -ne $imagen){try{$imagen.Dispose()}catch{}}}
}

function Obtener-Paginas-Word{
 param([object]$word,[string]$ruta)
 $documento=$null
 try{
  $documento=$word.Documents.Open($ruta,$false,$true)
  $documento.Repaginate()
  $paginas=$documento.ComputeStatistics(2)
  $documento.Close($false)
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($documento)|Out-Null
  $documento=$null
  if($paginas -lt 1){$paginas=1}
  return $paginas
 }catch{
  if($null -ne $documento){
   try{$documento.Close($false)}catch{}
   try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($documento)|Out-Null}catch{}
  }
  return 0
 }
}

function Obtener-Paginas-DOCXFallback{
 param([string]$ruta)
 $zip=$null
 try{
  Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
  $zip=[System.IO.Compression.ZipFile]::OpenRead($ruta)
  $entrada=$zip.GetEntry("docProps/app.xml")
  if($null -eq $entrada){return 0}
  $lector=New-Object System.IO.StreamReader($entrada.Open())
  $xmlTexto=$lector.ReadToEnd()
  $lector.Close()
  $xml=New-Object System.Xml.XmlDocument
  $xml.LoadXml($xmlTexto)
  $nodo=$xml.SelectSingleNode("//*[local-name()='Pages']")
  if($null -eq $nodo){return 0}
  $paginas=0
  if([int]::TryParse($nodo.InnerText,[ref]$paginas) -and $paginas -gt 0){return $paginas}
  return 0
 }catch{return 0}
 finally{if($null -ne $zip){try{$zip.Dispose()}catch{}}}
}

function Obtener-Paginas-PDF{
 param([string]$ruta)

 $script:PdfDiagnostic = ""

 if(!$script:PdfEngineReady){
  $script:PdfDiagnostic = "El motor PDF no esta listo.`r`n`r`n$script:PdfEngineMessage"
  return 0
 }

 $stream=$null

 try{
  if(!(Test-Path $ruta -PathType Leaf)){
   throw "El archivo no existe: $ruta"
  }

  $info=Get-Item $ruta -ErrorAction Stop
  if($info.Length -lt 5){
   throw "El archivo tiene un tamaño demasiado pequeño para ser un PDF ($($info.Length) bytes)."
  }

  $stream=[System.IO.File]::Open(
      $ruta,
      [System.IO.FileMode]::Open,
      [System.IO.FileAccess]::Read,
      [System.IO.FileShare]::ReadWrite
  )

  $headerBytes=New-Object byte[] 5
  [void]$stream.Read($headerBytes,0,5)
  $header=[System.Text.Encoding]::ASCII.GetString($headerBytes)
  $stream.Position=0

  $asmPdf=[PDFtoImage.Conversion].Assembly
  $asmSkia=[SkiaSharp.SKBitmap].Assembly

  $paginas=[PDFtoImage.Conversion]::GetPageCount($stream,$true,$null)

  if($paginas -lt 1){
   throw "PDFium devolvio $paginas paginas. Encabezado detectado: '$header'."
  }

  $script:PdfDiagnostic = @"
PDF OK

Archivo: $($info.Name)
Tamaño: $($info.Length) bytes
Encabezado: $header
Paginas detectadas: $paginas

PDFtoImage: $($asmPdf.FullName)
SkiaSharp: $($asmSkia.FullName)
"@

  return [int]$paginas
 }
 catch{
  $ex=$_.Exception
  $inner=$ex.InnerException
  $detalles=@()

  while($null -ne $inner){
   [void]$detalles.Add("INNER: $($inner.GetType().FullName)`r`n$($inner.Message)")
   $inner=$inner.InnerException
  }

  $asmPdfInfo="No disponible"
  $asmSkiaInfo="No disponible"
  try{$asmPdfInfo=[PDFtoImage.Conversion].Assembly.FullName}catch{}
  try{$asmSkiaInfo=[SkiaSharp.SKBitmap].Assembly.FullName}catch{}

  $headerTexto="No se pudo leer"
  try{
   if($null -ne $stream){
    $stream.Position=0
    $hb=New-Object byte[] 5
    [void]$stream.Read($hb,0,5)
    $headerTexto=[System.Text.Encoding]::ASCII.GetString($hb)
   }
  }catch{}

  $innerTexto=""
  if($detalles.Count -gt 0){
   $innerTexto="`r`n`r`nDetalles internos:`r`n"+($detalles -join "`r`n`r`n")
  }

  $script:PdfDiagnostic=@"
NO SE PUDO LEER EL PDF

Archivo: $ruta
Tamaño: $((Get-Item $ruta -ErrorAction SilentlyContinue).Length) bytes
Encabezado: $headerTexto

EXCEPCION:
$($ex.GetType().FullName)
$($ex.Message)

PDFtoImage cargado:
$asmPdfInfo

SkiaSharp cargado:
$asmSkiaInfo
$innerTexto
"@

  return 0
 }
 finally{
  if($null -ne $stream){
   try{$stream.Dispose()}catch{}
  }
 }
}
function Preparar-Archivo{
 param([System.IO.FileInfo]$archivo)
 $extension=$archivo.Extension.ToLower()
 $paginas=0
 if($extension-in @(".tif",".tiff")){$paginas=Obtener-Paginas-TIFF $archivo.FullName}
 elseif($extension-in @(".pdf")){$paginas=Obtener-Paginas-PDF $archivo.FullName}
 $archivo|Add-Member -NotePropertyName PageCount -NotePropertyValue $paginas -Force
 $archivo|Add-Member -NotePropertyName PrintRange -NotePropertyValue "" -Force
 $archivo|Add-Member -NotePropertyName PageError -NotePropertyValue $script:PdfDiagnostic -Force
 return $archivo
}

function Agregar-Archivo{
 param([string]$ruta)
 if(!(Test-Path $ruta -PathType Leaf)){return}
 $archivo=Get-Item $ruta
 $extension=$archivo.Extension.ToLower()
 if($extensionesPermitidas-notcontains $extension){return}
 $yaExiste=$archivos|Where-Object{$_.FullName-eq$archivo.FullName}
 if($null-ne $yaExiste){return}
 $preparado=Preparar-Archivo $archivo
 [void]$archivos.Add($preparado)
 if($extension -eq ".pdf" -and $preparado.PageCount -le 0){
  $mensaje=$preparado.PageError
  if([string]::IsNullOrWhiteSpace($mensaje)){$mensaje="No se obtuvo informacion adicional."}
  [System.Windows.Forms.MessageBox]::Show(
      $mensaje,
      "Diagnostico PDF - ECP",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
  )
 }
}

function Actualizar-Tabla{
 $tabla.Rows.Clear()
 for($i=0;$i-lt $archivos.Count;$i++){
  $archivo=$archivos[$i]
  $fila=$tabla.Rows.Add()
  $tabla.Rows[$fila].Cells["No"].Value="{0:D2}"-f($i+1)
  $tabla.Rows[$fila].Cells["Archivo"].Value=$archivo.Name
  $tabla.Rows[$fila].Cells["Paginas"].Value=if($archivo.PageCount-gt 0){$archivo.PageCount}else{"..."}
  $tabla.Rows[$fila].Cells["Rango"].Value=$archivo.PrintRange
 }
 $labelInfo.Text="Documentos en cola: $($archivos.Count)"
}

function Actualizar-Paginas-Word{
 $hayWord=$archivos|Where-Object{$_.Extension.ToLower()-in @(".doc",".docx")}
 if($null-eq $hayWord){return}
 $word=$null
 try{
  $labelEstado.Text="Estado: Calculando paginas de Word..."
  $labelEstado.ForeColor=$colorTextoSecundario
  Procesar-Eventos
  try{
   $word=New-Object -ComObject Word.Application
   $word.Visible=$false;$word.DisplayAlerts=0
  }catch{$word=$null}

  foreach($archivo in $hayWord){
   $paginas=0
   if($null -ne $word){
    try{$paginas=Obtener-Paginas-Word $word $archivo.FullName}catch{$paginas=0}
   }
   if($paginas -le 0 -and $archivo.Extension.ToLower() -eq ".docx"){
    $paginas=Obtener-Paginas-DOCXFallback $archivo.FullName
   }
   if($paginas -gt 0){
    $archivo.PageCount=$paginas
    $archivo.PageError=""
   }
  }
 }catch{}
 finally{
  if($null-ne $word){
   try{$word.Quit()}catch{}
   try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)|Out-Null}catch{}
  }
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
 }
 Actualizar-Tabla
}

function Parsear-Rango-Paginas{
 param([string]$rango,[int]$totalPaginas)
 if([string]::IsNullOrWhiteSpace($rango)){return @(1..$totalPaginas)}
 $resultado=New-Object System.Collections.Generic.List[int]
 foreach($parteOriginal in $rango.Split(",")){
  $parte=$parteOriginal.Trim()
  if([string]::IsNullOrWhiteSpace($parte)){throw "Hay una parte vacia en el rango."}
  if($parte-match '^(\d+)\s*-\s*(\d+)$'){
   $inicio=[int]$Matches[1];$fin=[int]$Matches[2]
   if($inicio-lt 1-or $fin-lt 1){throw "Las paginas deben comenzar en 1."}
   if($inicio-gt $fin){throw "El rango '$parte' esta invertido."}
   if($fin-gt $totalPaginas){throw "El rango '$parte' supera las $totalPaginas paginas."}
   for($pagina=$inicio;$pagina-le $fin;$pagina++){if(!$resultado.Contains($pagina)){[void]$resultado.Add($pagina)}}
   continue
  }
  if($parte-match '^\d+$'){
   $pagina=[int]$parte
   if($pagina-lt 1){throw "La pagina debe ser mayor o igual a 1."}
   if($pagina-gt $totalPaginas){throw "La pagina $pagina supera las $totalPaginas paginas."}
   if(!$resultado.Contains($pagina)){[void]$resultado.Add($pagina)}
   continue
  }
  throw "Formato invalido: '$parte'. Usa 2, 1-3 o 1-3,5,8-10."
 }
 return @($resultado)
}

function Obtener-Bloques-Rango{
 param([string]$rango,[int]$totalPaginas)
 if([string]::IsNullOrWhiteSpace($rango)){return @(@{Inicio=1;Fin=$totalPaginas})}
 $bloques=New-Object System.Collections.ArrayList
 foreach($parteOriginal in $rango.Split(",")){
  $parte=$parteOriginal.Trim()
  if($parte-match '^(\d+)\s*-\s*(\d+)$'){$inicio=[int]$Matches[1];$fin=[int]$Matches[2]}
  elseif($parte-match '^\d+$'){$inicio=[int]$parte;$fin=$inicio}
  else{throw "Formato invalido: '$parte'."}
  if($inicio-lt 1-or $fin-lt 1){throw "Las paginas deben comenzar en 1."}
  if($inicio-gt $fin){throw "El rango '$parte' esta invertido."}
  if($fin-gt $totalPaginas){throw "El rango '$parte' supera las $totalPaginas paginas."}
  [void]$bloques.Add(@{Inicio=$inicio;Fin=$fin})
 }
 return @($bloques)
}

function Validar-Rango{
 param([string]$rango,[int]$totalPaginas)
 if([string]::IsNullOrWhiteSpace($rango)){return @{Valido=$true;Mensaje=""}}
 if($totalPaginas-le 0){return @{Valido=$false;Mensaje="No se pudo determinar el numero de paginas."}}
 try{[void](Parsear-Rango-Paginas $rango $totalPaginas);return @{Valido=$true;Mensaje=""}}
 catch{return @{Valido=$false;Mensaje=$_.Exception.Message}}
}

$tabla.Add_CellEndEdit({
 param($sender,$e)
 if($e.RowIndex-lt 0-or $e.ColumnIndex-ne $tabla.Columns["Rango"].Index-or $e.RowIndex-ge $archivos.Count){return}
 $valor=$tabla.Rows[$e.RowIndex].Cells["Rango"].Value
 if($null-eq $valor){$valor=""}
 $valor=$valor.ToString().Trim()
 $archivo=$archivos[$e.RowIndex]
 $resultado=Validar-Rango $valor $archivo.PageCount
 $archivo.PrintRange=$valor
 if(!$resultado.Valido){
  $tabla.Rows[$e.RowIndex].Cells["Rango"].Style.ForeColor=$colorRojo
  $tabla.Rows[$e.RowIndex].Cells["Rango"].ToolTipText=$resultado.Mensaje
  $labelEstado.Text="Rango invalido: $($resultado.Mensaje)";$labelEstado.ForeColor=$colorRojo
 }else{
  $tabla.Rows[$e.RowIndex].Cells["Rango"].Style.ForeColor=$colorTexto
  $tabla.Rows[$e.RowIndex].Cells["Rango"].ToolTipText=""
  $labelEstado.Text="Estado: Lista para imprimir";$labelEstado.ForeColor=$colorVerde
 }
})

function Crear-Boton{
 param([string]$texto,[int]$x,[int]$ancho)
 $b=New-Object System.Windows.Forms.Button
 $b.Text=$texto;$b.Size=New-Object System.Drawing.Size($ancho,40);$b.Location=New-Object System.Drawing.Point($x,580)
 $b.BackColor=$colorControl;$b.ForeColor=$colorTexto;$b.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat;$b.FlatAppearance.BorderColor=$colorBorde
 $form.Controls.Add($b);return $b
}

$btnAgregar=Crear-Boton "Agregar archivos" 25 150
$btnEliminar=Crear-Boton "Eliminar" 185 110
$btnSubir=Crear-Boton "Subir" 305 100
$btnBajar=Crear-Boton "Bajar" 415 100
$btnVaciar=Crear-Boton "Vaciar cola" 525 110

$btnAgregar.Add_Click({
 $dialogo=New-Object System.Windows.Forms.OpenFileDialog
 $dialogo.Multiselect=$true
 $dialogo.Filter="Documentos|*.pdf;*.doc;*.docx;*.tif;*.tiff|PDF|*.pdf|Word|*.doc;*.docx|TIFF|*.tif;*.tiff"
 $dialogo.Title="Seleccionar documentos"
 if($dialogo.ShowDialog()-eq [System.Windows.Forms.DialogResult]::OK){
  foreach($ruta in $dialogo.FileNames){Agregar-Archivo $ruta}
  Actualizar-Tabla;Actualizar-Paginas-Word
 }
})

$tabla.AllowDrop=$true
$tabla.Add_DragEnter({
 if($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)){$_.Effect=[System.Windows.Forms.DragDropEffects]::Copy}else{$_.Effect=[System.Windows.Forms.DragDropEffects]::None}
})
$tabla.Add_DragDrop({
 if($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)){
  $rutas=$_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
  foreach($ruta in $rutas){Agregar-Archivo $ruta}
  Actualizar-Tabla;Actualizar-Paginas-Word
 }
})

$btnEliminar.Add_Click({
 $fila=$tabla.CurrentRow;if($null-eq $fila){return};$indice=$fila.Index
 if($indice-ge 0-and $indice-lt $archivos.Count){
  $archivos.RemoveAt($indice);Actualizar-Tabla
  if($archivos.Count-gt 0){$nuevoIndice=[Math]::Min($indice,$archivos.Count-1);$tabla.ClearSelection();$tabla.Rows[$nuevoIndice].Selected=$true;$tabla.CurrentCell=$tabla.Rows[$nuevoIndice].Cells["Archivo"]}
 }
})

$btnSubir.Add_Click({
 $fila=$tabla.CurrentRow;if($null-eq $fila){return};$indice=$fila.Index
 if($indice-le 0){return}
 $temporal=$archivos[$indice];$archivos[$indice]=$archivos[$indice-1];$archivos[$indice-1]=$temporal
 Actualizar-Tabla;$nuevoIndice=$indice-1;$tabla.ClearSelection();$tabla.Rows[$nuevoIndice].Selected=$true;$tabla.CurrentCell=$tabla.Rows[$nuevoIndice].Cells["Archivo"]
 $tabla.FirstDisplayedScrollingRowIndex=[Math]::Max(0,$nuevoIndice)
})

$btnBajar.Add_Click({
 $fila=$tabla.CurrentRow;if($null-eq $fila){return};$indice=$fila.Index
 if($indice-lt 0-or $indice-ge ($archivos.Count-1)){return}
 $temporal=$archivos[$indice];$archivos[$indice]=$archivos[$indice+1];$archivos[$indice+1]=$temporal
 Actualizar-Tabla;$nuevoIndice=$indice+1;$tabla.ClearSelection();$tabla.Rows[$nuevoIndice].Selected=$true;$tabla.CurrentCell=$tabla.Rows[$nuevoIndice].Cells["Archivo"]
 $tabla.FirstDisplayedScrollingRowIndex=[Math]::Max(0,$nuevoIndice)
})

$btnVaciar.Add_Click({
 if($archivos.Count-eq 0){return}
 $respuesta=[System.Windows.Forms.MessageBox]::Show("Deseas eliminar todos los documentos?","Vaciar cola",
 [System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
 if($respuesta-eq [System.Windows.Forms.DialogResult]::Yes){$archivos.Clear();Actualizar-Tabla;$labelEstado.Text="Estado: Lista para imprimir";$labelEstado.ForeColor=$colorVerde}
})

function Obtener-Impresora-Predeterminada{
 $printer=Get-CimInstance Win32_Printer|Where-Object{$_.Default-eq $true}|Select-Object -First 1
 if($null-eq $printer){return $null};return $printer.Name
}
function Cambiar-Impresora-Predeterminada{
 param([string]$nombreImpresora)
 $network=New-Object -ComObject WScript.Network
 $network.SetDefaultPrinter($nombreImpresora)
 Start-Sleep -Milliseconds 700
}

function Convertir-DOCX-a-PDF{
 param([string]$rutaDocx)

 $tempPdf = Join-Path $env:TEMP ("ColaImpresion_DOCX_" + [guid]::NewGuid().ToString("N") + ".pdf")
 $tempVbs = Join-Path $env:TEMP ("ColaImpresion_DOCX_" + [guid]::NewGuid().ToString("N") + ".vbs")

 $docxEsc = $rutaDocx.Replace("\","\\").Replace('"','""')
 $pdfEsc = $tempPdf.Replace("\","\\").Replace('"','""')

 $vbs = @"
Set Word = CreateObject("Word.Application")
Word.Visible = False
Set Doc = Word.Documents.Open("$docxEsc", False, True)
Doc.ExportAsFixedFormat "$pdfEsc", 17, False
Doc.Close False
Word.Quit
Set Doc = Nothing
Set Word = Nothing
"@

 try{
  [System.IO.File]::WriteAllText($tempVbs,$vbs,[System.Text.Encoding]::Default)
  $p=Start-Process "cscript.exe" -ArgumentList @("//nologo","`"$tempVbs`"") -Wait -PassThru -WindowStyle Hidden
  if($p.ExitCode -ne 0 -or !(Test-Path $tempPdf)){
   throw "Word no pudo convertir el documento a PDF. Verifica que Microsoft Word este instalado en este equipo."
  }
  return $tempPdf
 }finally{
  if(Test-Path $tempVbs){Remove-Item $tempVbs -Force -ErrorAction SilentlyContinue}
 }
}

# ------------------------------------------------------------
# FIX APLICADO: se agrego el parametro $nombreImpresora.
# Antes, Imprimir-Word llamaba a Imprimir-PDF usando $nombreImpresora
# como si fuera una variable global, pero en PowerShell el scope de
# una funcion es lexico (segun donde se DEFINIO), no dinamico (segun
# quien la llamo). Como $nombreImpresora solo existia dentro del
# scriptblock del boton Imprimir, dentro de Imprimir-Word esa variable
# era $null. Resultado: PrinterSettings.PrinterName = $null, la
# impresora se marcaba como invalida, y CUALQUIER Word/DOCX fallaba
# siempre, en cualquier equipo. Ahora se recibe explicitamente.
# ------------------------------------------------------------
function Imprimir-Word{
 param([object]$word,[string]$ruta,[string]$nombreImpresora,[string]$rango,[int]$totalPaginas)

 $pdfTemporal=$null
 try{
  $labelEstado.Text="Convirtiendo Word a PDF..."
  $labelEstado.ForeColor=$colorTextoSecundario
  Procesar-Eventos

  $pdfTemporal=Convertir-DOCX-a-PDF $ruta

  # Reutilizamos exactamente el motor PDF ya probado.
  [void](Imprimir-PDF $pdfTemporal $nombreImpresora $rango $totalPaginas)

  return $true
 }finally{
  if($null -ne $pdfTemporal -and (Test-Path $pdfTemporal)){
   Remove-Item $pdfTemporal -Force -ErrorAction SilentlyContinue
  }
 }
}

function Imprimir-TIFF{
 param([string]$ruta,[string]$nombreImpresora,[string]$rango,[int]$totalPaginas)

 $imagen=$null

 try{
  $imagen=[System.Drawing.Image]::FromFile($ruta)
  $dimension=[System.Drawing.Imaging.FrameDimension]::Page
  $frames=$imagen.GetFrameCount($dimension)

  if($frames-lt 1){$frames=1}

  $bloques=Obtener-Bloques-Rango $rango $totalPaginas

  foreach($bloque in $bloques){

   $printDocument=New-Object System.Drawing.Printing.PrintDocument

   try{
    $printDocument.PrinterSettings.PrinterName=$nombreImpresora

    if(!$printDocument.PrinterSettings.IsValid){
     throw "La impresora seleccionada no esta disponible: $nombreImpresora"
    }

    $printDocument.DocumentName=[System.IO.Path]::GetFileName($ruta)

    $printDocument.DefaultPageSettings.Margins=
        New-Object System.Drawing.Printing.Margins(0,0,0,0)

    $paginasBloque =
        New-Object System.Collections.Generic.List[int]

    for($pagina=$bloque.Inicio;$pagina-le $bloque.Fin;$pagina++){
     [void]$paginasBloque.Add($pagina)
    }

    $script:paginasTIFF=$paginasBloque
    $script:indicePaginaTIFF=0

    $printDocument.Add_PrintPage({
     param($sender,$e)

     if($script:indicePaginaTIFF-ge $script:paginasTIFF.Count){
      $e.HasMorePages=$false
      return
     }

     $pagina=$script:paginasTIFF[$script:indicePaginaTIFF]

     $imagen.SelectActiveFrame(
         $dimension,
         $pagina-1
     )

     $dpiX=[double]$imagen.HorizontalResolution
     $dpiY=[double]$imagen.VerticalResolution

     if($dpiX-le 1){$dpiX=200}
     if($dpiY-le 1){$dpiY=200}

     $area=$e.PageSettings.PrintableArea

     $anchoFisico=([double]$imagen.Width/$dpiX)*100
     $altoFisico=([double]$imagen.Height/$dpiY)*100

     $anchoDisponible=[double]$area.Width
     $altoDisponible=[double]$area.Height

     $escala=1.0

     if($anchoFisico-gt $anchoDisponible){
      $escala=[Math]::Min(
          $escala,
          $anchoDisponible/$anchoFisico
      )
     }

     if($altoFisico-gt $altoDisponible){
      $escala=[Math]::Min(
          $escala,
          $altoDisponible/$altoFisico
      )
     }

     $anchoFinal=[int]($anchoFisico*$escala)
     $altoFinal=[int]($altoFisico*$escala)

     $x=[int](
         $area.X+
         (($anchoDisponible-$anchoFinal)/2)
     )

     $y=[int](
         $area.Y+
         (($altoDisponible-$altoFinal)/2)
     )

     $e.Graphics.Clear(
         [System.Drawing.Color]::White
     )

     $e.Graphics.InterpolationMode=
         [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

     $e.Graphics.SmoothingMode=
         [System.Drawing.Drawing2D.SmoothingMode]::HighQuality

     $e.Graphics.PixelOffsetMode=
         [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

     $e.Graphics.DrawImage(
         $imagen,
         $x,
         $y,
         $anchoFinal,
         $altoFinal
     )

     $script:indicePaginaTIFF++

     $e.HasMorePages=
         ($script:indicePaginaTIFF-lt $script:paginasTIFF.Count)
    })

    $printDocument.Print()

    Start-Sleep -Milliseconds 500
    Procesar-Eventos
   }
   finally{
    if($null-ne $printDocument){
     try{$printDocument.Dispose()}catch{}
    }
   }
  }

  return $true
 }
 catch{
  throw $_
 }
 finally{
  if($null-ne $imagen){
   try{$imagen.Dispose()}catch{}
  }
 }
}

function Imprimir-PDF{
 param(
  [string]$ruta,
  [string]$nombreImpresora,
  [string]$rango,
  [int]$totalPaginas
 )

 if(!$script:PdfEngineReady){
  throw "No fue posible inicializar el motor PDF integrado. $script:PdfEngineMessage"
 }

 $bloques=Obtener-Bloques-Rango $rango $totalPaginas

 foreach($bloque in $bloques){

  $printDocument=$null
  $stream=$null

  try{

   $printDocument=New-Object System.Drawing.Printing.PrintDocument
   $printDocument.PrinterSettings.PrinterName=$nombreImpresora

   if(!$printDocument.PrinterSettings.IsValid){
    throw "La impresora seleccionada no esta disponible: $nombreImpresora"
   }

   $printDocument.DocumentName=[System.IO.Path]::GetFileName($ruta)

   $printDocument.DefaultPageSettings.Margins=
       New-Object System.Drawing.Printing.Margins(0,0,0,0)

   $stream=[System.IO.File]::Open(
       $ruta,
       [System.IO.FileMode]::Open,
       [System.IO.FileAccess]::Read,
       [System.IO.FileShare]::Read
   )

   $script:pdfStreamActual=$stream

   $script:pdfPaginasBloque=
       New-Object System.Collections.Generic.List[int]

   for(
       $pagina=$bloque.Inicio;
       $pagina-le $bloque.Fin;
       $pagina++
   ){
    [void]$script:pdfPaginasBloque.Add($pagina)
   }

   $script:pdfIndicePagina=0

   $printDocument.Add_PrintPage({

    param($sender,$e)

    if(
        $script:pdfIndicePagina-ge
        $script:pdfPaginasBloque.Count
    ){
     $e.HasMorePages=$false
     return
    }

    $paginaPdf=
        $script:pdfPaginasBloque[
            $script:pdfIndicePagina
        ] - 1

    $bitmap=$null
    $pngStream=$null
    $imagen=$null

    try{

     $opciones=
         New-Object PDFtoImage.RenderOptions

     $bitmap=
         [PDFtoImage.Conversion]::ToImage(
             $script:pdfStreamActual,
             $paginaPdf,
             $true,
             $null,
             $opciones
         )

     if($null -eq $bitmap){
      throw "PDFium no pudo renderizar la pagina $($paginaPdf+1)."
     }

     $pngStream=
         New-Object System.IO.MemoryStream

     $ok=
         $bitmap.Encode(
             $pngStream,
             [SkiaSharp.SKEncodedImageFormat]::Png,
             100
         )

     if(!$ok){
      throw "No fue posible convertir la pagina PDF a imagen."
     }

     $pngStream.Position=0

     $imagen=
         [System.Drawing.Image]::FromStream(
             $pngStream
         )

     $area=$e.PageSettings.PrintableArea

     $dpi=300.0

     $anchoFisico=
         ([double]$imagen.Width/$dpi)*100

     $altoFisico=
         ([double]$imagen.Height/$dpi)*100

     $anchoDisponible=[double]$area.Width
     $altoDisponible=[double]$area.Height

     $escala=1.0

     if($anchoFisico-gt $anchoDisponible){
      $escala=[Math]::Min(
          $escala,
          $anchoDisponible/$anchoFisico
      )
     }

     if($altoFisico-gt $altoDisponible){
      $escala=[Math]::Min(
          $escala,
          $altoDisponible/$altoFisico
      )
     }

     $anchoFinal=[int]($anchoFisico*$escala)
     $altoFinal=[int]($altoFisico*$escala)

     $x=[int](
         $area.X+
         (($anchoDisponible-$anchoFinal)/2)
     )

     $y=[int](
         $area.Y+
         (($altoDisponible-$altoFinal)/2)
     )

     $e.Graphics.Clear(
         [System.Drawing.Color]::White
     )

     $e.Graphics.InterpolationMode=
         [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

     $e.Graphics.SmoothingMode=
         [System.Drawing.Drawing2D.SmoothingMode]::HighQuality

     $e.Graphics.PixelOffsetMode=
         [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

     $e.Graphics.DrawImage(
         $imagen,
         $x,
         $y,
         $anchoFinal,
         $altoFinal
     )

     $script:pdfIndicePagina++

     $e.HasMorePages=
         ($script:pdfIndicePagina-lt
          $script:pdfPaginasBloque.Count)
    }
    finally{

     if($null-ne $imagen){
      try{$imagen.Dispose()}catch{}
     }

     if($null-ne $pngStream){
      try{$pngStream.Dispose()}catch{}
     }

     if($null-ne $bitmap){
      try{$bitmap.Dispose()}catch{}
     }
    }
   })

   $labelEstado.Text=
       "Imprimiendo PDF: paginas $($bloque.Inicio)-$($bloque.Fin)"

   Procesar-Eventos

   $printDocument.Print()

   Start-Sleep -Milliseconds 500
   Procesar-Eventos
  }
  finally{

   if($null-ne $stream){
    try{$stream.Dispose()}catch{}
   }

   if($null-ne $printDocument){
    try{$printDocument.Dispose()}catch{}
   }
  }
 }

 return $true
}


# ============================================================
# APOYO AL PROYECTO
# ============================================================
# Todos los textos, banco y CLABE se leen de $script:ConfigECP,
# ya resuelto y verificado (o su respaldo local) desde el inicio
# del script. Aqui ya no se descarga nada de internet.

function Mostrar-Apoyo-Proyecto {

    try {

        $config = $script:ConfigECP

        $formApoyo = New-Object System.Windows.Forms.Form
        $formApoyo.Text = "Acomplétame para una coquita"
        $formApoyo.Size = New-Object System.Drawing.Size(760,460)
        $formApoyo.StartPosition = "CenterParent"
        $formApoyo.BackColor = $colorFondo
        $formApoyo.ForeColor = $colorTexto
        $formApoyo.Font = New-Object System.Drawing.Font("Segoe UI",10)
        $formApoyo.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $formApoyo.MaximizeBox = $false
        $formApoyo.MinimizeBox = $false

        $lblTituloApoyo = New-Object System.Windows.Forms.Label
        $lblTituloApoyo.Text = [string]$config.titulo
        $lblTituloApoyo.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
        $lblTituloApoyo.Location = New-Object System.Drawing.Point(28,24)
        $lblTituloApoyo.Size = New-Object System.Drawing.Size(430,30)
        $lblTituloApoyo.ForeColor = $colorTexto
        $formApoyo.Controls.Add($lblTituloApoyo)

        $lblMensajeAmigos = New-Object System.Windows.Forms.Label
        $lblMensajeAmigos.Text = [string]$config.mensaje
        $lblMensajeAmigos.Location = New-Object System.Drawing.Point(28,60)
        $lblMensajeAmigos.Size = New-Object System.Drawing.Size(450,40)
        $lblMensajeAmigos.ForeColor = $colorTexto
        $formApoyo.Controls.Add($lblMensajeAmigos)

        $lblApoyo = New-Object System.Windows.Forms.Label
        $lblApoyo.Text = "Si quieres apoyar al desarrollador:"
        $lblApoyo.Font = New-Object System.Drawing.Font("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
        $lblApoyo.Location = New-Object System.Drawing.Point(28,110)
        $lblApoyo.Size = New-Object System.Drawing.Size(430,30)
        $lblApoyo.ForeColor = $colorAzul
        $formApoyo.Controls.Add($lblApoyo)

        $lblClabeTitulo = New-Object System.Windows.Forms.Label
        $lblClabeTitulo.Text = "CLABE interbancaria"
        $lblClabeTitulo.Location = New-Object System.Drawing.Point(28,148)
        $lblClabeTitulo.AutoSize = $true
        $lblClabeTitulo.ForeColor = $colorTextoSecundario
        $formApoyo.Controls.Add($lblClabeTitulo)

        $txtClabe = New-Object System.Windows.Forms.TextBox
        $txtClabe.Text = [string]$config.clabe
        $txtClabe.Location = New-Object System.Drawing.Point(28,173)
        $txtClabe.Size = New-Object System.Drawing.Size(300,32)
        $txtClabe.ReadOnly = $true
        $txtClabe.BackColor = $colorControl
        $txtClabe.ForeColor = $colorTexto
        $txtClabe.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $formApoyo.Controls.Add($txtClabe)

        $btnCopiarClabe = New-Object System.Windows.Forms.Button
        $btnCopiarClabe.Text = "Copiar"
        $btnCopiarClabe.Location = New-Object System.Drawing.Point(338,173)
        $btnCopiarClabe.Size = New-Object System.Drawing.Size(90,32)
        $btnCopiarClabe.BackColor = $colorAzul
        $btnCopiarClabe.ForeColor = [System.Drawing.Color]::White
        $btnCopiarClabe.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $formApoyo.Controls.Add($btnCopiarClabe)

        $btnCopiarClabe.Add_Click({
            try {
                [System.Windows.Forms.Clipboard]::SetText($txtClabe.Text)
                [System.Windows.Forms.MessageBox]::Show(
                    "CLABE copiada al portapapeles.",
                    "Apoyo ECP",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            } catch {}
        })

        $lblBeneficiarioTitulo = New-Object System.Windows.Forms.Label
        $lblBeneficiarioTitulo.Text = "Beneficiario:"
        $lblBeneficiarioTitulo.Location = New-Object System.Drawing.Point(28,220)
        $lblBeneficiarioTitulo.AutoSize = $true
        $lblBeneficiarioTitulo.ForeColor = $colorTextoSecundario
        $formApoyo.Controls.Add($lblBeneficiarioTitulo)

        $lblBeneficiario = New-Object System.Windows.Forms.Label
        $lblBeneficiario.Text = [string]$config.beneficiario
        $lblBeneficiario.Location = New-Object System.Drawing.Point(28,244)
        $lblBeneficiario.Size = New-Object System.Drawing.Size(430,25)
        $lblBeneficiario.ForeColor = $colorTexto
        $formApoyo.Controls.Add($lblBeneficiario)

        $lblBancoTitulo = New-Object System.Windows.Forms.Label
        $lblBancoTitulo.Text = "Banco:"
        $lblBancoTitulo.Location = New-Object System.Drawing.Point(28,278)
        $lblBancoTitulo.AutoSize = $true
        $lblBancoTitulo.ForeColor = $colorTextoSecundario
        $formApoyo.Controls.Add($lblBancoTitulo)

        $lblBanco = New-Object System.Windows.Forms.Label
        $lblBanco.Text = [string]$config.banco
        $lblBanco.Location = New-Object System.Drawing.Point(28,302)
        $lblBanco.AutoSize = $true
        $lblBanco.ForeColor = $colorTexto
        $formApoyo.Controls.Add($lblBanco)

        $lblNoObligatorio = New-Object System.Windows.Forms.Label
        $lblNoObligatorio.Text = [string]$config.aviso
        $lblNoObligatorio.Location = New-Object System.Drawing.Point(28,340)
        $lblNoObligatorio.Size = New-Object System.Drawing.Size(430,28)
        $lblNoObligatorio.ForeColor = $colorTextoSecundario
        $formApoyo.Controls.Add($lblNoObligatorio)

        if ($script:ConfigAdvertenciaECP -eq "FirmaInvalida") {
            $lblModoRespaldo = New-Object System.Windows.Forms.Label
            $lblModoRespaldo.Text = "Mostrando datos de respaldo local (no se pudo verificar la configuración remota)."
            $lblModoRespaldo.Location = New-Object System.Drawing.Point(28,368)
            $lblModoRespaldo.Size = New-Object System.Drawing.Size(430,32)
            $lblModoRespaldo.ForeColor = $colorNaranja
            $formApoyo.Controls.Add($lblModoRespaldo)
        }

        $panelCoquita = New-Object System.Windows.Forms.Panel
        $panelCoquita.Location = New-Object System.Drawing.Point(485,18)
        $panelCoquita.Size = New-Object System.Drawing.Size(245,400)
        $panelCoquita.BackColor = $colorFondo
        $formApoyo.Controls.Add($panelCoquita)

        $rutaPerro2 = Join-Path $basePath "perro2.png"
        if(Test-Path $rutaPerro2){
            try{
                $picPerro2 = New-Object System.Windows.Forms.PictureBox
                $picPerro2.Location = New-Object System.Drawing.Point(0,0)
                $picPerro2.Size = New-Object System.Drawing.Size(245,325)
                $picPerro2.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
                $picPerro2.BackColor = $colorFondo
                $picPerro2.Image = [System.Drawing.Image]::FromFile($rutaPerro2)
                $panelCoquita.Controls.Add($picPerro2)
            }catch{}
        }

        $lblVersion = New-Object System.Windows.Forms.Label
        $lblVersion.Text = "Verifica que esta sea la versión oficial ECP: $([string]$config.version)"
        $lblVersion.Location = New-Object System.Drawing.Point(28,390)
        $lblVersion.Size = New-Object System.Drawing.Size(520,25)
        $lblVersion.ForeColor = $colorTextoSecundario
        $formApoyo.Controls.Add($lblVersion)

        [void]$formApoyo.ShowDialog($form)
        $formApoyo.Dispose()

    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "No fue posible mostrar la ventana de apoyo.`r`n`r`n" +
            "Esto no afecta la impresión.`r`n`r`n" +
            "Detalle: " + $_.Exception.Message,
            "Apoyo ECP",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
}

$btnApoyo = New-Object System.Windows.Forms.Button
$btnApoyo.Text = "Acomplétame para una coquita"
$btnApoyo.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$btnApoyo.Size = New-Object System.Drawing.Size(170,40)
$btnApoyo.Location = New-Object System.Drawing.Point(640,580)
$btnApoyo.BackColor = [System.Drawing.Color]::FromArgb(32,32,32)
$btnApoyo.ForeColor = $colorTexto
$btnApoyo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnApoyo.FlatAppearance.BorderColor = $colorBorde
$btnApoyo.FlatAppearance.BorderSize = 1
$form.Controls.Add($btnApoyo)

# Interruptor remoto "activo": si el desarrollador lo apaga en el JSON
# firmado, el boton se deshabilita (solo afecta donaciones, la impresion
# sigue funcionando con normalidad).
if (-not [bool]$script:ConfigECP.activo) {
    $btnApoyo.Enabled = $false
    $btnApoyo.Text = "Apoyo no disponible por ahora"
}

$btnApoyo.Add_Click({
    Mostrar-Apoyo-Proyecto
})


$btnImprimir=New-Object System.Windows.Forms.Button
$btnImprimir.Text="IMPRIMIR"
$btnImprimir.Font=New-Object System.Drawing.Font("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
$btnImprimir.Size=New-Object System.Drawing.Size(170,40)
$btnImprimir.Location=New-Object System.Drawing.Point(830,580)
$btnImprimir.BackColor=[System.Drawing.Color]::FromArgb(22,32,48)
$btnImprimir.ForeColor=$colorTexto
$btnImprimir.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$btnImprimir.FlatAppearance.BorderColor=$colorAzul
$btnImprimir.FlatAppearance.BorderSize=2
$form.Controls.Add($btnImprimir)

$btnImprimir.Add_Click({
 if($archivos.Count-eq 0){
  [System.Windows.Forms.MessageBox]::Show("No hay documentos en la cola.","Cola vacia",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information);return
 }
 if($comboImpresoras.SelectedIndex-lt 0){
  [System.Windows.Forms.MessageBox]::Show("Selecciona una impresora.","Impresora",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning);return
 }

 for($i=0;$i-lt $archivos.Count;$i++){
  $archivo=$archivos[$i];$ext=$archivo.Extension.ToLower()
  if($archivo.PageCount-le 0-and $ext-in @(".pdf",".doc",".docx",".tif",".tiff")){
   $detalle=$archivo.PageError
   if([string]::IsNullOrWhiteSpace($detalle)){$detalle="No se obtuvo informacion adicional."}
   [System.Windows.Forms.MessageBox]::Show("No se pudo determinar el numero de paginas de '$($archivo.Name)'.`r`n`r`n$detalle","Pagina desconocida",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning);return
  }
  $resultado=Validar-Rango $archivo.PrintRange $archivo.PageCount
  if(!$resultado.Valido){
   $tabla.ClearSelection();$tabla.Rows[$i].Selected=$true;$tabla.CurrentCell=$tabla.Rows[$i].Cells["Rango"]
   [System.Windows.Forms.MessageBox]::Show("Hay un rango invalido en:`n`n$($archivo.Name)`n`n$($resultado.Mensaje)","Rango invalido",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning);return
  }
 }

 $nombreImpresora=$comboImpresoras.SelectedItem.ToString()
 $respuesta=[System.Windows.Forms.MessageBox]::Show("Se enviaran $($archivos.Count) documentos a:`n`n$nombreImpresora`n`nEl orden sera exactamente el mostrado en la cola.`n`nDeseas comenzar?","Iniciar impresion",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
 if($respuesta-ne [System.Windows.Forms.DialogResult]::Yes){return}

 $impresoraOriginal=Obtener-Impresora-Predeterminada
 $btnAgregar.Enabled=$false;$btnEliminar.Enabled=$false;$btnSubir.Enabled=$false;$btnBajar.Enabled=$false;$btnVaciar.Enabled=$false;$btnImprimir.Enabled=$false;$comboImpresoras.Enabled=$false
 $errores=New-Object System.Collections.ArrayList;$word=$null

 try{
  $labelEstado.Text="Seleccionando impresora...";$labelEstado.ForeColor=$colorTextoSecundario;Procesar-Eventos
  Cambiar-Impresora-Predeterminada $nombreImpresora

  # DOC/DOCX se convierten a PDF mediante VBScript; no necesitamos controlar
  # Word desde PowerShell COM durante la impresion.

  for($i=0;$i-lt $archivos.Count;$i++){
   $archivo=$archivos[$i];$numero=$i+1;$extension=$archivo.Extension.ToLower()
   $tabla.ClearSelection();$tabla.Rows[$i].Selected=$true;$tabla.CurrentCell=$tabla.Rows[$i].Cells["Archivo"]
   $labelEstado.Text="Imprimiendo $numero de $($archivos.Count): $($archivo.Name)";$labelEstado.ForeColor=$colorTextoSecundario;Procesar-Eventos
   try{
    if($extension-in @(".doc",".docx")){[void](Imprimir-Word $word $archivo.FullName $nombreImpresora $archivo.PrintRange $archivo.PageCount)}
    elseif($extension-in @(".tif",".tiff")){[void](Imprimir-TIFF $archivo.FullName $nombreImpresora $archivo.PrintRange $archivo.PageCount)}
    elseif($extension-eq ".pdf"){[void](Imprimir-PDF $archivo.FullName $nombreImpresora $archivo.PrintRange $archivo.PageCount)}
   }catch{[void]$errores.Add("$($archivo.Name): $($_.Exception.Message)")}
   Procesar-Eventos
  }
 }catch{
  [System.Windows.Forms.MessageBox]::Show("Ocurrio un error general:`n`n$($_.Exception.Message)","Error de impresion",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
 }
 finally{
  if($null-ne $word){
   try{$word.Quit()}catch{}
   try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)|Out-Null}catch{}
   $word=$null
  }
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
  if(![string]::IsNullOrWhiteSpace($impresoraOriginal)){try{Cambiar-Impresora-Predeterminada $impresoraOriginal}catch{}}
  $btnAgregar.Enabled=$true;$btnEliminar.Enabled=$true;$btnSubir.Enabled=$true;$btnBajar.Enabled=$true;$btnVaciar.Enabled=$true;$btnImprimir.Enabled=$true;$comboImpresoras.Enabled=$true
 }

 if($errores.Count-eq 0){
  $labelEstado.Text="Estado: Impresion terminada";$labelEstado.ForeColor=$colorVerde
  [System.Windows.Forms.MessageBox]::Show("Los $($archivos.Count) documentos fueron procesados correctamente.","Impresion terminada",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
 }else{
  $labelEstado.Text="Estado: Terminado con advertencias";$labelEstado.ForeColor=$colorNaranja
  $mensaje="La cola termino, pero hubo problemas con:`n`n"+(($errores|ForEach-Object{"- $_`n"})-join "")
  [System.Windows.Forms.MessageBox]::Show($mensaje,"Advertencia",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
 }
})

$tabla.Add_CellDoubleClick({
 param($sender,$e)
 if($e.RowIndex-ge 0-and $e.RowIndex-lt $archivos.Count){Start-Process $archivos[$e.RowIndex].FullName}
})

Actualizar-Tabla
[void]$form.ShowDialog()
