#Requires -Version 5.1
<#
.SYNOPSIS
    Diagnóstico de red para Windows. Registra, con fecha y hora, la latencia y las
    pérdidas hacia el módem y hacia internet, el DNS, el estado del Wi-Fi, el tráfico
    que genera la propia PC y pruebas de velocidad periódicas con latencia bajo carga.
    Al terminar escribe un resumen que separa fallas de la casa y fallas del proveedor.

.DESCRIPTION
    No requiere permisos de administrador ni instala nada.
    Archivos que genera (en el Escritorio, carpeta Diagnostico_Red_AAAAMMDD_HHMM):
      muestras.csv           una fila cada IntervaloSegundos
      pruebas_velocidad.csv  una fila por prueba de velocidad
      info_inicial.txt       configuración de red, Wi-Fi y ruta (tracert) al inicio
      resumen.txt            estadísticas e interpretación al terminar
    y un .zip con todo, listo para enviarlo.

.PARAMETER Minutos
    Duración del registro. 0 = hasta que presiones Ctrl+C. Predeterminado: 60.

.PARAMETER IntervaloSegundos
    Cada cuántos segundos se toma una muestra. Mínimo 2. Predeterminado: 5.

.PARAMETER VelocidadCadaMinutos
    Cada cuántos minutos se hace una prueba de velocidad. 0 = no hacer pruebas.
    Predeterminado: 15. Cada prueba descarga y sube unos 8 segundos de datos por sentido.

.PARAMETER PlanMbps
    Velocidad de bajada contratada, si la conoces, para compararla. Opcional.

.PARAMETER Destinos
    Direcciones de internet a las que se hace ping. Predeterminado: 8.8.8.8 y 1.1.1.1.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\diagnostico_red.ps1
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\diagnostico_red.ps1 -Minutos 180 -PlanMbps 100
#>
[CmdletBinding()]
param(
    [int]$Minutos = 60,
    [int]$IntervaloSegundos = 5,
    [int]$VelocidadCadaMinutos = 15,
    [int]$PlanMbps = 0,
    [string[]]$Destinos = @('8.8.8.8', '1.1.1.1'),
    [string]$Carpeta = [Environment]::GetFolderPath('Desktop'),
    [string]$NombreDns = 'claude.ai',
    [string]$UrlBajada = 'https://speed.cloudflare.com/__down?bytes=',
    [string]$UrlSubida = 'https://speed.cloudflare.com/__up',
    [switch]$SinTracert
)

$ErrorActionPreference = 'Continue'

# Con -File, "8.8.8.8,1.1.1.1" llega como un solo texto.
$Destinos = @($Destinos | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($Destinos.Count -eq 0) { $Destinos = @('8.8.8.8', '1.1.1.1') }
if ($IntervaloSegundos -lt 2) { $IntervaloSegundos = 2 }
if (-not $Carpeta) { $Carpeta = (Get-Location).Path }

$script:CsvEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$script:Utf8Bom = New-Object System.Text.UTF8Encoding($true)

# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

function Get-ZonaUtc([datetime]$Fecha) {
    $off = [TimeZoneInfo]::Local.GetUtcOffset($Fecha)
    $signo = if ($off.Ticks -lt 0) { '-' } else { '+' }
    '{0}{1:00}:{2:00}' -f $signo, [math]::Abs($off.Hours), [math]::Abs($off.Minutes)
}

function ConvertTo-Numero([string]$Texto) {
    if (-not $Texto) { return $null }
    $m = [regex]::Match($Texto, '-?\d+([.,]\d+)?')
    if (-not $m.Success) { return $null }
    [double]::Parse($m.Value.Replace(',', '.'), [Globalization.CultureInfo]::InvariantCulture)
}

function Get-Estadistica($Valores) {
    $o = @($Valores | Where-Object { $null -ne $_ } | ForEach-Object { [double]$_ } | Sort-Object)
    if ($o.Count -eq 0) { return $null }
    $n = $o.Count
    [pscustomobject]@{
        N       = $n
        Min     = $o[0]
        Prom    = [math]::Round(($o | Measure-Object -Average).Average, 1)
        Mediana = $o[[int][math]::Floor(($n - 1) / 2)]
        P95     = $o[[int][math]::Ceiling(0.95 * $n) - 1]
        Max     = $o[$n - 1]
    }
}

function Save-Fila([string]$Ruta, $Objeto) {
    $Objeto | Export-Csv -Path $Ruta -Append -NoTypeInformation -UseCulture -Encoding $script:CsvEncoding
}

function Format-Ms($Resultado) {
    if ($null -eq $Resultado) { return 'n/d' }
    if ($null -ne $Resultado.Ms) { return "$($Resultado.Ms) ms" }
    $Resultado.Estado
}

# ---------------------------------------------------------------------------
# Mediciones
# ---------------------------------------------------------------------------

# Interfaz con puerta de enlace IPv4. Si hay cable y Wi-Fi, Windows prefiere el cable.
function Get-InterfazActiva {
    $candidatas = foreach ($nic in [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
        if ($nic.OperationalStatus.ToString() -ne 'Up') { continue }
        $tipo = $nic.NetworkInterfaceType.ToString()
        if ($tipo -eq 'Loopback' -or $tipo -eq 'Tunnel') { continue }
        try { $props = $nic.GetIPProperties() } catch { continue }
        $gw = $props.GatewayAddresses |
            Where-Object { $_.Address.AddressFamily.ToString() -eq 'InterNetwork' -and $_.Address.ToString() -ne '0.0.0.0' } |
            Select-Object -First 1
        if (-not $gw) { continue }
        $dns = $props.DnsAddresses |
            Where-Object { $_.AddressFamily.ToString() -eq 'InterNetwork' } |
            Select-Object -First 1
        [pscustomobject]@{
            Nic     = $nic
            Tipo    = $tipo
            Gateway = $gw.Address.ToString()
            Dns     = if ($dns) { $dns.ToString() } else { $null }
        }
    }
    $candidatas | Sort-Object { if ($_.Tipo -eq 'Wireless80211') { 1 } else { 0 } } | Select-Object -First 1
}

function Get-NombreConexion($Interfaz) {
    if (-not $Interfaz) { return 'Sin conexion' }
    if ($Interfaz.Tipo -eq 'Wireless80211') { return 'Wi-Fi' }
    if ($Interfaz.Tipo -match 'Ethernet') { return 'Cable' }
    $Interfaz.Tipo
}

# Pings en paralelo. Devuelve @{ ip = @{ Ms; Estado } }.
function Invoke-Pings([string[]]$Ips, [int]$TimeoutMs = 1000) {
    $tareas = @{}
    $pings = @()
    foreach ($ip in ($Ips | Where-Object { $_ } | Select-Object -Unique)) {
        $p = New-Object System.Net.NetworkInformation.Ping
        $pings += $p
        try { $tareas[$ip] = $p.SendPingAsync($ip, $TimeoutMs) } catch { $tareas[$ip] = $null }
    }
    $res = @{}
    foreach ($ip in @($tareas.Keys)) {
        $t = $tareas[$ip]
        $r = [pscustomobject]@{ Ms = $null; Estado = 'ERROR_LOCAL' }
        if ($t) {
            try { [void]$t.Wait($TimeoutMs + 1500) } catch { }
            if ($t.IsCompleted -and -not $t.IsFaulted -and -not $t.IsCanceled) {
                $reply = $t.Result
                switch ($reply.Status.ToString()) {
                    'Success'                       { $r.Ms = [int]$reply.RoundtripTime; $r.Estado = 'OK' }
                    'TimedOut'                      { $r.Estado = 'TIEMPO_AGOTADO' }
                    'DestinationHostUnreachable'    { $r.Estado = 'INALCANZABLE' }
                    'DestinationNetworkUnreachable' { $r.Estado = 'RED_INALCANZABLE' }
                    default                         { $r.Estado = $reply.Status.ToString().ToUpperInvariant() }
                }
            } elseif (-not $t.IsCompleted) {
                $r.Estado = 'TIEMPO_AGOTADO'
            }
            # Si la tarea falló (sin red local), queda ERROR_LOCAL: equivale al "Error general" de ping.exe.
        }
        $res[$ip] = $r
    }
    foreach ($p in $pings) { $p.Dispose() }
    $res
}

# Consulta DNS directa por UDP al servidor del adaptador. No usa la caché de Windows.
function Test-Dns([string]$Servidor, [string]$Nombre, [int]$TimeoutMs = 2000) {
    if (-not $Servidor) { return [pscustomobject]@{ Ms = $null; Estado = 'SIN_SERVIDOR' } }
    $udp = New-Object System.Net.Sockets.UdpClient
    try {
        $udp.Client.ReceiveTimeout = $TimeoutMs
        $id = Get-Random -Minimum 0 -Maximum 65535
        $pkt = New-Object System.Collections.Generic.List[byte]
        $pkt.AddRange([byte[]]@(($id -shr 8), ($id -band 0xFF), 0x01, 0x00, 0x00, 0x01, 0, 0, 0, 0, 0, 0))
        foreach ($etiqueta in $Nombre.Split('.')) {
            $b = [Text.Encoding]::ASCII.GetBytes($etiqueta)
            $pkt.Add([byte]$b.Length)
            $pkt.AddRange($b)
        }
        $pkt.AddRange([byte[]]@(0, 0, 1, 0, 1))
        $datos = $pkt.ToArray()
        $sw = [Diagnostics.Stopwatch]::StartNew()
        [void]$udp.Send($datos, $datos.Length, $Servidor, 53)
        $remoto = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        $resp = $udp.Receive([ref]$remoto)
        $sw.Stop()
        if ($resp.Length -ge 4 -and $resp[0] -eq $datos[0] -and $resp[1] -eq $datos[1]) {
            $rcode = $resp[3] -band 0x0F
            if ($rcode -eq 0) { return [pscustomobject]@{ Ms = [int]$sw.ElapsedMilliseconds; Estado = 'OK' } }
            return [pscustomobject]@{ Ms = $null; Estado = "RCODE_$rcode" }
        }
        [pscustomobject]@{ Ms = $null; Estado = 'RESPUESTA_INVALIDA' }
    } catch {
        [pscustomobject]@{ Ms = $null; Estado = 'SIN_RESPUESTA' }
    } finally {
        $udp.Close()
    }
}

# Lee "netsh wlan show interfaces" en español o inglés.
function Get-InfoWifi {
    $info = [pscustomobject]@{
        Estado = $null; Ssid = $null; Bssid = $null; Banda = $null; Canal = $null
        Senal = $null; Rssi = $null; Radio = $null; Rx = $null; Tx = $null; Nota = $null
    }
    try { $salida = (& netsh.exe wlan show interfaces 2>&1 | Out-String) } catch { $info.Nota = 'netsh no disponible'; return $info }

    $bloques = New-Object System.Collections.Generic.List[hashtable]
    $actual = $null
    foreach ($linea in ($salida -split "`r?`n")) {
        if ($linea -match '^\s*(.+?)\s+:\s(.*)$') {
            $clave = $Matches[1].Trim().ToLowerInvariant()
            $valor = $Matches[2].Trim()
            if ($clave -match '^(nombre|name)$') { $actual = @{}; $bloques.Add($actual) }
            if ($null -ne $actual -and -not $actual.ContainsKey($clave)) { $actual[$clave] = $valor }
        }
    }
    if ($bloques.Count -eq 0) {
        if ($salida -match 'ubicaci|location') { $info.Nota = 'SIN_PERMISO_UBICACION' } else { $info.Nota = 'SIN_DATOS_WIFI' }
        return $info
    }

    function Get-Campo($Bloque, [string]$Patron) {
        foreach ($k in $Bloque.Keys) { if ($k -match $Patron) { return $Bloque[$k] } }
        $null
    }

    $bloque = $bloques | Where-Object { (Get-Campo $_ '^(estado|state)$') -match '^(conectado|connected)$' } | Select-Object -First 1
    if (-not $bloque) { $bloque = $bloques[0] }

    $info.Estado = Get-Campo $bloque '^(estado|state)$'
    $info.Ssid   = Get-Campo $bloque '^ssid$'
    $info.Bssid  = Get-Campo $bloque '^(ap )?bssid$'
    $info.Radio  = Get-Campo $bloque '^(tipo de radio|radio type)$'
    $info.Banda  = Get-Campo $bloque '^(banda|band)$'
    $info.Canal  = ConvertTo-Numero (Get-Campo $bloque '^(canal|channel)$')
    $info.Senal  = ConvertTo-Numero (Get-Campo $bloque '^(se.{1,2}al|signal)$')
    $info.Rssi   = ConvertTo-Numero (Get-Campo $bloque '^rssi$')
    $info.Rx     = ConvertTo-Numero (Get-Campo $bloque '(recepci|receive)')
    $info.Tx     = ConvertTo-Numero (Get-Campo $bloque '(transmisi|transmit)')
    if (-not $info.Banda -and $info.Canal) {
        $info.Banda = if ($info.Canal -le 14) { '2.4 GHz (inferida)' } else { '5 GHz (inferida)' }
    }
    $info
}

# Qué programas tienen conexiones abiertas a internet (cuenta conexiones, no bytes).
function Get-ProcesosConConexiones {
    try {
        $con = Get-NetTCPConnection -State Established -ErrorAction Stop |
            Where-Object { $_.RemoteAddress -notmatch '^(127\.|::1$|0\.0\.0\.0$|::$)' }
    } catch { return '' }
    $nombres = @{}
    Get-Process | ForEach-Object { $nombres[$_.Id] = $_.ProcessName }
    $conteo = @{}
    foreach ($c in $con) {
        $n = $nombres[[int]$c.OwningProcess]
        if (-not $n) { $n = "PID$($c.OwningProcess)" }
        $conteo[$n] = 1 + [int]$conteo[$n]
    }
    ($conteo.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5 |
        ForEach-Object { "$($_.Key)($($_.Value))" }) -join '; '
}

# ---------------------------------------------------------------------------
# Prueba de velocidad con latencia bajo carga
# ---------------------------------------------------------------------------

function Initialize-Http {
    Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        [Net.ServicePointManager]::Expect100Continue = $false
    } catch { }
    $script:Http = New-Object System.Net.Http.HttpClient
    $script:Http.Timeout = [TimeSpan]::FromSeconds(90)
    $script:Http.DefaultRequestHeaders.ExpectContinue = $false
    [void]$script:Http.DefaultRequestHeaders.UserAgent.TryParseAdd('DiagnosticoRed/1.0')
}

function Measure-Transferencia([string]$Direccion, [int]$Bytes, [string]$Gateway, [string]$DestinoInternet) {
    $latModem = New-Object System.Collections.Generic.List[double]
    $latInternet = New-Object System.Collections.Generic.List[double]
    $perdidos = 0
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        if ($Direccion -eq 'Bajada') {
            $tarea = $script:Http.GetByteArrayAsync("$UrlBajada$Bytes")
        } else {
            $contenido = New-Object System.Net.Http.ByteArrayContent(, (New-Object byte[] $Bytes))
            $tarea = $script:Http.PostAsync($UrlSubida, $contenido)
        }
    } catch {
        return [pscustomobject]@{ Ok = $false; Mbps = $null; Bytes = 0; Segundos = 0; LatModem = $latModem; LatInternet = $latInternet; Perdidos = 0; Error = $_.Exception.Message }
    }
    while (-not $tarea.IsCompleted) {
        $p = Invoke-Pings @($Gateway, $DestinoInternet) 800
        foreach ($par in @(@($Gateway, $latModem), @($DestinoInternet, $latInternet))) {
            $r = if ($par[0]) { $p[$par[0]] } else { $null }
            if ($r -and $null -ne $r.Ms) { $par[1].Add($r.Ms) } elseif ($r) { $perdidos++ }
        }
        if (-not $tarea.IsCompleted) { Start-Sleep -Milliseconds 250 }
    }
    $sw.Stop()
    $seg = $sw.Elapsed.TotalSeconds
    $error_ = $null
    $recibidos = 0
    if ($tarea.IsFaulted -or $tarea.IsCanceled) {
        $error_ = if ($tarea.Exception) { $tarea.Exception.GetBaseException().Message } else { 'Cancelada (tiempo agotado)' }
    } elseif ($Direccion -eq 'Bajada') {
        $recibidos = $tarea.Result.Length
    } else {
        if ($tarea.Result.IsSuccessStatusCode) { $recibidos = $Bytes } else { $error_ = "HTTP $([int]$tarea.Result.StatusCode)" }
        $tarea.Result.Dispose()
    }
    $mbps = if (-not $error_ -and $seg -gt 0) { [math]::Round($recibidos * 8 / $seg / 1e6, 1) } else { $null }
    [pscustomobject]@{
        Ok = (-not $error_); Mbps = $mbps; Bytes = $recibidos; Segundos = [math]::Round($seg, 1)
        LatModem = $latModem; LatInternet = $latInternet; Perdidos = $perdidos; Error = $error_
    }
}

function Get-TamanoObjetivo($Mbps, [int]$Min, [int]$Max) {
    if (-not $Mbps) { return $Min }
    $b = [int][math]::Min([double]$Max, [math]::Max([double]$Min, $Mbps * 1e6 / 8 * 8))   # ~8 segundos
    $b
}

function Invoke-PruebaVelocidad($Interfaz, [string]$Gateway) {
    $destino = $Destinos[0]
    $ahora = Get-Date
    Write-Host ("{0}  Prueba de velocidad en curso (unos 30 s)..." -f $ahora.ToString('HH:mm:ss')) -ForegroundColor Cyan

    # Latencia en reposo justo antes de la prueba.
    $repM = @(); $repI = @()
    1..5 | ForEach-Object {
        $p = Invoke-Pings @($Gateway, $destino) 1000
        if ($Gateway -and $p[$Gateway] -and $null -ne $p[$Gateway].Ms) { $repM += $p[$Gateway].Ms }
        if ($p[$destino] -and $null -ne $p[$destino].Ms) { $repI += $p[$destino].Ms }
        Start-Sleep -Milliseconds 200
    }

    # Abre la conexión TLS antes de medir para que no cuente en el tiempo.
    try { [void]$script:Http.GetByteArrayAsync("${UrlBajada}0").Wait(15000) } catch { }

    $sondaB = Measure-Transferencia 'Bajada' 2000000 $Gateway $destino
    $baj = if ($sondaB.Ok) { Measure-Transferencia 'Bajada' (Get-TamanoObjetivo $sondaB.Mbps 5000000 100000000) $Gateway $destino } else { $sondaB }
    $sondaS = Measure-Transferencia 'Subida' 1000000 $Gateway $destino
    $sub = if ($sondaS.Ok) { Measure-Transferencia 'Subida' (Get-TamanoObjetivo $sondaS.Mbps 2000000 50000000) $Gateway $destino } else { $sondaS }

    $eBM = Get-Estadistica $baj.LatModem; $eBI = Get-Estadistica $baj.LatInternet
    $eSM = Get-Estadistica $sub.LatModem; $eSI = Get-Estadistica $sub.LatInternet
    $eRM = Get-Estadistica $repM; $eRI = Get-Estadistica $repI
    $errores = @($baj.Error, $sub.Error | Where-Object { $_ }) -join ' / '

    $fila = [pscustomobject][ordered]@{
        FechaHora                     = $ahora.ToString('yyyy-MM-dd HH:mm:ss')
        UTC                           = Get-ZonaUtc $ahora
        Conexion                      = Get-NombreConexion $Interfaz
        Bajada_Mbps                   = $baj.Mbps
        Subida_Mbps                   = $sub.Mbps
        MB_bajada                     = [math]::Round($baj.Bytes / 1e6, 1)
        MB_subida                     = [math]::Round($sub.Bytes / 1e6, 1)
        Modem_reposo_ms               = if ($eRM) { $eRM.Mediana } else { $null }
        Internet_reposo_ms            = if ($eRI) { $eRI.Mediana } else { $null }
        Modem_carga_bajada_ms         = if ($eBM) { $eBM.Mediana } else { $null }
        Internet_carga_bajada_ms      = if ($eBI) { $eBI.Mediana } else { $null }
        Internet_carga_bajada_max_ms  = if ($eBI) { $eBI.Max } else { $null }
        Modem_carga_subida_ms         = if ($eSM) { $eSM.Mediana } else { $null }
        Internet_carga_subida_ms      = if ($eSI) { $eSI.Mediana } else { $null }
        Internet_carga_subida_max_ms  = if ($eSI) { $eSI.Max } else { $null }
        Pings_perdidos_en_carga       = $baj.Perdidos + $sub.Perdidos
        Error                         = $errores
    }
    $color = if ($errores) { 'Red' } else { 'Cyan' }
    Write-Host ("{0}  Velocidad: baja {1} Mbps, sube {2} Mbps | latencia a internet: reposo {3} ms, bajando {4} ms, subiendo {5} ms {6}" -f `
        $ahora.ToString('HH:mm:ss'), $fila.Bajada_Mbps, $fila.Subida_Mbps, $fila.Internet_reposo_ms,
        $fila.Internet_carga_bajada_ms, $fila.Internet_carga_subida_ms, $errores) -ForegroundColor $color
    $fila
}

# ---------------------------------------------------------------------------
# Resumen
# ---------------------------------------------------------------------------

function Write-Resumen {
    $L = New-Object System.Collections.Generic.List[string]
    $termino = Get-Date
    $L.Add('RESUMEN DEL DIAGNÓSTICO DE RED')
    $L.Add(("Inicio: {0}   Fin: {1}   Zona horaria: UTC{2} ({3})" -f $inicio.ToString('yyyy-MM-dd HH:mm:ss'), $termino.ToString('yyyy-MM-dd HH:mm:ss'), (Get-ZonaUtc $termino), [TimeZoneInfo]::Local.Id))
    $L.Add(("Muestras: {0} (cada {1} s)   Pruebas de velocidad: {2}" -f $muestras.Count, $IntervaloSegundos, $pruebas.Count))
    if ($muestras.Count -eq 0) {
        $L.Add('No se tomaron muestras.')
        return $L
    }
    $conexiones = @($muestras | Group-Object Conexion | ForEach-Object { "{0}: {1} muestras" -f $_.Name, $_.Count })
    $L.Add('Tipo de conexión: ' + ($conexiones -join '; '))
    $L.Add('')

    $interp = New-Object System.Collections.Generic.List[string]

    # Latencia y pérdidas
    $L.Add('LATENCIA Y PÉRDIDAS (ms: mínimo / promedio / mediana / p95 / máximo)')
    $objetivos = @(, @('Modem', "Módem ($(($muestras | Where-Object { $_.Modem_IP } | Select-Object -Last 1).Modem_IP))"))
    foreach ($d in $Destinos) { $objetivos += , @($d, $d) }
    $perdidas = @{}
    foreach ($o in $objetivos) {
        $col = $o[0]
        $validas = @($muestras | Where-Object { $_."$($col)_estado" })
        $fallas = @($validas | Where-Object { $_."$($col)_estado" -ne 'OK' }).Count
        $pct = if ($validas.Count) { [math]::Round(100 * $fallas / $validas.Count, 2) } else { 0 }
        $perdidas[$col] = $pct
        $e = Get-Estadistica ($validas | ForEach-Object { $_."$($col)_ms" })
        if ($e) {
            $L.Add(("  {0,-22} pérdidas {1,6}% ({2} de {3})   {4} / {5} / {6} / {7} / {8}" -f $o[1], $pct, $fallas, $validas.Count, $e.Min, $e.Prom, $e.Mediana, $e.P95, $e.Max))
        } else {
            $L.Add(("  {0,-22} pérdidas {1,6}% ({2} de {3})   sin respuestas" -f $o[1], $pct, $fallas, $validas.Count))
        }
        if ($col -eq 'Modem' -and $e -and $e.P95 -gt 30) {
            $interp.Add("La latencia hacia el módem es alta en el 5% peor de las muestras (p95 = $($e.P95) ms; con cable suele ser de 1 a 2 ms y con Wi-Fi sano, de pocos ms). Indicio de problema entre la PC y el módem: Wi-Fi débil o con interferencia, o la PC ocupada.")
        }
    }
    $internetPerd = ($Destinos | ForEach-Object { $perdidas[$_] } | Measure-Object -Minimum).Minimum
    if ($perdidas['Modem'] -ge 1) {
        $interp.Add("Hay $($perdidas['Modem'])% de pérdidas hacia el módem. Eso ocurre dentro de la casa (Wi-Fi, cable o el propio módem), no en la red del proveedor.")
    } elseif ($internetPerd -ge 1) {
        $interp.Add("El módem responde bien pero hay $internetPerd% de pérdidas hacia internet. Eso apunta al módem (lado internet) o al proveedor.")
    }

    # DNS
    $dnsVal = @($muestras | Where-Object { $_.DNS_estado -and $_.DNS_estado -ne 'SIN_SERVIDOR' })
    if ($dnsVal.Count) {
        $dnsF = @($dnsVal | Where-Object { $_.DNS_estado -ne 'OK' })
        $dnsFconInternet = @($dnsF | Where-Object { $_.Internet_OK -eq 'SI' }).Count
        $e = Get-Estadistica ($dnsVal | ForEach-Object { $_.DNS_ms })
        $txt = if ($e) { "mediana $($e.Mediana) ms, p95 $($e.P95) ms" } else { 'sin respuestas' }
        $L.Add(("  DNS ({0}, consulta {1})   fallas {2} de {3}; {4}" -f ($dnsVal | Select-Object -Last 1).DNS_servidor, $NombreDns, $dnsF.Count, $dnsVal.Count, $txt))
        if ($dnsFconInternet -gt 0) {
            $interp.Add("El DNS falló $dnsFconInternet veces mientras el ping a internet sí respondía. En esos momentos las páginas no cargan aunque haya conexión. Probar otro DNS (por ejemplo 1.1.1.1 u 8.8.8.8) es una solución barata.")
        }
    }
    $L.Add('')

    # Cortes
    $L.Add('CORTES (ningún destino de internet respondió)')
    $eventos = New-Object System.Collections.Generic.List[object]
    $ev = $null
    foreach ($m in $muestras) {
        if ($m.Internet_OK -eq 'NO') {
            if (-not $ev) { $ev = @{ Inicio = $m.FechaHora; Filas = 0; ModemFalla = 0; Regreso = $null } }
            $ev['Filas'] = $ev['Filas'] + 1
            if ($m.Modem_estado -ne 'OK') { $ev['ModemFalla'] = $ev['ModemFalla'] + 1 }
        } elseif ($ev) {
            $ev['Regreso'] = $m.FechaHora
            $eventos.Add([pscustomobject]$ev)
            $ev = $null
        }
    }
    if ($ev) { $eventos.Add([pscustomobject]$ev) }
    if ($eventos.Count -eq 0) { $L.Add('  Ninguno.') }
    $locales = 0; $externos = 0
    foreach ($e in $eventos) {
        if ($e.ModemFalla -eq $e.Filas) { $tipo = 'LOCAL: la PC tampoco llegaba al módem (Wi-Fi/cable, o módem reiniciándose)'; $locales++ }
        elseif ($e.ModemFalla -eq 0) { $tipo = 'DESPUÉS DEL MÓDEM: el módem respondía pero internet no (módem lado internet o proveedor)'; $externos++ }
        else { $tipo = 'MIXTO' }
        $ini = [datetime]::ParseExact($e.Inicio, 'yyyy-MM-dd HH:mm:ss', $null)
        if ($e.Regreso) {
            $dur = [int]([datetime]::ParseExact($e.Regreso, 'yyyy-MM-dd HH:mm:ss', $null) - $ini).TotalSeconds
            $L.Add(("  {0} -> {1}  ({2} s)  {3}" -f $e.Inicio, $e.Regreso, $dur, $tipo))
        } else {
            $L.Add(("  {0} -> seguía caído al terminar  {1}" -f $e.Inicio, $tipo))
        }
    }
    if ($externos -gt 0) { $interp.Add("Hubo $externos corte(s) con el módem respondiendo: la falla estuvo del módem hacia afuera. Esta es la evidencia útil para un reporte al proveedor (anota fecha y hora).") }
    if ($locales -gt 0) { $interp.Add("Hubo $locales corte(s) en los que la PC tampoco llegaba al módem: la falla estuvo dentro de la casa. Si en ese momento reiniciaste el módem, es lo esperado.") }
    $L.Add('')

    # Tráfico de la PC
    $L.Add('TRÁFICO QUE GENERÓ ESTA PC (medido en el adaptador, sin contar las pruebas de velocidad)')
    $eB = Get-Estadistica ($muestras | ForEach-Object { $_.PC_baja_Mbps })
    $eS = Get-Estadistica ($muestras | ForEach-Object { $_.PC_sube_Mbps })
    if ($eB -and $eS) {
        $L.Add(("  Bajada: promedio {0} Mbps, p95 {1}, máximo {2}" -f $eB.Prom, $eB.P95, $eB.Max))
        $L.Add(("  Subida: promedio {0} Mbps, p95 {1}, máximo {2}" -f $eS.Prom, $eS.P95, $eS.Max))
        $conSubida = @($muestras | Where-Object { $null -ne $_.PC_sube_Mbps -and $_.PC_sube_Mbps -gt 2 })
        $pctSubida = [math]::Round(100 * $conSubida.Count / $eS.N, 1)
        $L.Add("  Muestras con la PC subiendo más de 2 Mbps: $pctSubida%")
        $top = @{}
        foreach ($m in ($muestras | Where-Object { $_.Procesos_con_conexiones })) {
            foreach ($parte in ($m.Procesos_con_conexiones -split ';\s*')) {
                if ($parte -match '^(.+)\((\d+)\)$') { $top[$Matches[1]] = 1 + [int]$top[$Matches[1]] }
            }
        }
        $topTxt = ($top.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 6 | ForEach-Object { "$($_.Key) (en $($_.Value) capturas)" }) -join ', '
        if ($topTxt) { $L.Add("  Programas con más conexiones abiertas durante el tráfico alto: $topTxt") }
        if ($pctSubida -ge 20) {
            $interp.Add("La PC estuvo subiendo más de 2 Mbps en el $pctSubida% del tiempo. Una subida sostenida satura el enlace y dispara la latencia. Revisa respaldos y sincronización (OneDrive, Google Drive, fotos, etc.). Programas candidatos: $topTxt. Para ver bytes por programa usa el Monitor de recursos (resmon), pestaña Red.")
        }
    } else {
        $L.Add('  Sin datos.')
    }
    $L.Add('')

    # Wi-Fi
    $wifi = @($muestras | Where-Object { $null -ne $_.WiFi_senal_pct })
    if ($wifi.Count) {
        $L.Add('WI-FI')
        $eW = Get-Estadistica ($wifi | ForEach-Object { $_.WiFi_senal_pct })
        $eR = Get-Estadistica ($wifi | ForEach-Object { $_.WiFi_vel_rx_Mbps })
        $bandas = ($wifi | Where-Object { $_.WiFi_banda } | Group-Object WiFi_banda | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', '
        $canales = ($wifi | Where-Object { $_.WiFi_canal } | Group-Object WiFi_canal | ForEach-Object { $_.Name }) -join ', '
        $bssids = @($wifi | Where-Object { $_.WiFi_BSSID } | Select-Object -ExpandProperty WiFi_BSSID -Unique).Count
        $L.Add(("  Red: {0}   Bandas: {1}   Canales: {2}   Puntos de acceso distintos (BSSID): {3}" -f ($wifi | Select-Object -Last 1).WiFi_SSID, $bandas, $canales, $bssids))
        $L.Add(("  Señal: mínimo {0}%, promedio {1}%   Velocidad de enlace (recepción): mediana {2} Mbps, mínimo {3}" -f $eW.Min, $eW.Prom, $(if ($eR) { $eR.Mediana } else { 'n/d' }), $(if ($eR) { $eR.Min } else { 'n/d' })))
        if ($eW.Prom -lt 60) { $interp.Add("La señal Wi-Fi promedio fue de $($eW.Prom)%. Es baja: acerca la PC al módem, usa cable o agrega un punto de acceso.") }
        if ($eR -and $eR.Mediana -lt 100) { $interp.Add("La velocidad de enlace Wi-Fi (entre la PC y el módem) tuvo una mediana de $($eR.Mediana) Mbps. En la práctica se obtiene bastante menos que eso, así que el Wi-Fi puede estar limitando la velocidad.") }
        $L.Add('')
    } elseif (@($muestras | Where-Object { $_.Nota -match 'UBICACION' }).Count) {
        $L.Add('WI-FI: sin datos porque Windows exige el permiso de ubicación para leerlos.')
        $L.Add('')
    }

    # Velocidad
    $L.Add('PRUEBAS DE VELOCIDAD (servidor: ' + ([uri]$UrlSubida).Host + ')')
    $ok = @($pruebas | Where-Object { $null -ne $_.Bajada_Mbps })
    foreach ($p in $pruebas) {
        $L.Add(("  {0} [{1}] baja {2} Mbps, sube {3} Mbps | ms a internet: reposo {4}, bajando {5}, subiendo {6} | ms al módem: reposo {7}, bajando {8}, subiendo {9} {10}" -f `
            $p.FechaHora, $p.Conexion, $p.Bajada_Mbps, $p.Subida_Mbps, $p.Internet_reposo_ms, $p.Internet_carga_bajada_ms,
            $p.Internet_carga_subida_ms, $p.Modem_reposo_ms, $p.Modem_carga_bajada_ms, $p.Modem_carga_subida_ms, $p.Error))
    }
    if ($pruebas.Count -eq 0) { $L.Add('  No se hicieron pruebas.') }
    if ($ok.Count) {
        $eBaj = Get-Estadistica ($ok | ForEach-Object { $_.Bajada_Mbps })
        $eSub = Get-Estadistica ($ok | ForEach-Object { $_.Subida_Mbps })
        $L.Add(("  Bajada: mediana {0} Mbps (mín {1}, máx {2})   Subida: mediana {3} Mbps" -f $eBaj.Mediana, $eBaj.Min, $eBaj.Max, $(if ($eSub) { $eSub.Mediana } else { 'n/d' })))
        if ($PlanMbps -gt 0) {
            $pct = [math]::Round(100 * $eBaj.Mediana / $PlanMbps)
            $L.Add("  Respecto al plan declarado de $PlanMbps Mbps: $pct% (mediana)")
        }

        # Bufferbloat: cuánto sube la latencia con carga y dónde se forma la cola.
        foreach ($sentido in @(@('bajar', 'Internet_carga_bajada_ms', 'Modem_carga_bajada_ms'), @('subir', 'Internet_carga_subida_ms', 'Modem_carga_subida_ms'))) {
            $difI = Get-Estadistica ($ok | Where-Object { $null -ne $_.($sentido[1]) -and $null -ne $_.Internet_reposo_ms } | ForEach-Object { $_.($sentido[1]) - $_.Internet_reposo_ms })
            $difM = Get-Estadistica ($ok | Where-Object { $null -ne $_.($sentido[2]) -and $null -ne $_.Modem_reposo_ms } | ForEach-Object { $_.($sentido[2]) - $_.Modem_reposo_ms })
            if ($difI -and $difI.Mediana -gt 100) {
                if ($difM -and $difM.Mediana -gt 50) {
                    $interp.Add("Al $($sentido[0]) datos, la latencia a internet sube $($difI.Mediana) ms y la latencia al módem también sube $($difM.Mediana) ms. La cola se forma entre la PC y el módem: apunta al Wi-Fi o a la propia PC.")
                } else {
                    $interp.Add("Al $($sentido[0]) datos, la latencia a internet sube $($difI.Mediana) ms pero la del módem casi no cambia. La cola se forma del módem hacia afuera: módem (lado internet) o proveedor. Un router con SQM (fq_codel/CAKE) lo mitiga si el problema es de saturación.")
                }
            }
        }
        $soloWifi = @($ok | Where-Object { $_.Conexion -ne 'Cable' }).Count -eq $ok.Count
        if ($soloWifi) { $interp.Add('Todas las pruebas de velocidad se hicieron sin cable. Para atribuir una velocidad baja al proveedor falta repetirlas con la PC conectada por cable al módem y sin otros aparatos usando internet.') }
    }
    $L.Add('')

    $L.Add('INTERPRETACIÓN AUTOMÁTICA (indicios, no conclusiones definitivas)')
    if ($interp.Count -eq 0) { $interp.Add('No se detectaron anomalías con los umbrales usados. Compara las velocidades con tu plan contratado.') }
    $i = 1
    foreach ($t in $interp) { $L.Add("  $i. $t"); $i++ }
    $L.Add('')
    $L.Add('Límites: el script no conoce tu plan contratado; los pings miden la ruta hasta 8.8.8.8/1.1.1.1, no hasta cada sitio; las pruebas de velocidad usan un solo servidor y una sola conexión, así que en planes muy rápidos pueden quedar por debajo de lo real.')
    $L
}

# ---------------------------------------------------------------------------
# Inicio
# ---------------------------------------------------------------------------

$inicio = Get-Date
$dir = Join-Path $Carpeta ('Diagnostico_Red_' + $inicio.ToString('yyyyMMdd_HHmm'))
New-Item -ItemType Directory -Path $dir -Force | Out-Null
$csvMuestras = Join-Path $dir 'muestras.csv'
$csvVelocidad = Join-Path $dir 'pruebas_velocidad.csv'
$txtInfo = Join-Path $dir 'info_inicial.txt'
$txtResumen = Join-Path $dir 'resumen.txt'
$fin = if ($Minutos -gt 0) { $inicio.AddMinutes($Minutos) } else { [datetime]::MaxValue }

Write-Host ''
Write-Host 'DIAGNÓSTICO DE RED' -ForegroundColor White
Write-Host ("Resultados en: {0}" -f $dir)
Write-Host ("Duración: {0}. Muestra cada {1} s. Prueba de velocidad: {2}." -f `
    $(if ($Minutos -gt 0) { "$Minutos min" } else { 'hasta Ctrl+C' }), $IntervaloSegundos,
    $(if ($VelocidadCadaMinutos -gt 0) { "cada $VelocidadCadaMinutos min" } else { 'desactivada' }))
Write-Host 'Puedes terminar antes con Ctrl+C; el resumen se genera igual.' -ForegroundColor DarkGray
Write-Host ''

# Información inicial
$info = New-Object System.Text.StringBuilder
function Add-Info([string]$Titulo, [string]$Texto) {
    [void]$info.AppendLine("===== $Titulo =====")
    [void]$info.AppendLine($Texto.TrimEnd())
    [void]$info.AppendLine('')
}
Add-Info 'Fecha y hora de inicio' ("{0} (UTC{1}, {2})" -f $inicio.ToString('yyyy-MM-dd HH:mm:ss'), (Get-ZonaUtc $inicio), [TimeZoneInfo]::Local.Id)
$so = try { (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Caption } catch { '' }
Add-Info 'Sistema' ("{0} {1} | PowerShell {2}" -f $so, [Environment]::OSVersion.Version, $PSVersionTable.PSVersion)
Add-Info 'Parámetros' ("Minutos={0} IntervaloSegundos={1} VelocidadCadaMinutos={2} PlanMbps={3} Destinos={4} DNS={5}" -f $Minutos, $IntervaloSegundos, $VelocidadCadaMinutos, $PlanMbps, ($Destinos -join ','), $NombreDns)
$act0 = Get-InterfazActiva
if ($act0) {
    Add-Info 'Interfaz activa' ("{0} ({1}) | tipo {2} | puerta de enlace (módem) {3} | DNS {4} | velocidad del enlace {5} Mbps" -f `
        $act0.Nic.Name, $act0.Nic.Description, (Get-NombreConexion $act0), $act0.Gateway, $act0.Dns, [math]::Round($act0.Nic.Speed / 1e6))
} else {
    Add-Info 'Interfaz activa' 'Ninguna interfaz con puerta de enlace: la PC no tiene conexión con el módem en este momento.'
}
try { Add-Info 'Configuración IP' (Get-NetIPConfiguration -ErrorAction Stop | Out-String -Width 200) } catch { }
try { Add-Info 'netsh wlan show interfaces' (& netsh.exe wlan show interfaces 2>&1 | Out-String) } catch { }
$patrones = 'OneDrive', 'GoogleDriveFS', 'Dropbox', 'iCloud*', 'MEGAsync', 'pCloud*', 'Box', 'steam', 'EpicGamesLauncher',
            'Battle.net', 'qbittorrent', 'uTorrent', 'BitTorrent', 'transmission*', 'obs64', 'Zoom', 'ms-teams', 'Teams'
$sync = @(Get-Process | Where-Object { $n = $_.ProcessName; @($patrones | Where-Object { $n -like $_ }).Count -gt 0 } |
          Select-Object -ExpandProperty ProcessName -Unique)
Add-Info 'Programas de sincronización, descargas o videollamada abiertos' $(if ($sync.Count) { $sync -join ', ' } else { 'Ninguno de la lista conocida.' })
if (-not $SinTracert) {
    Write-Host 'Trazando la ruta hacia internet (tracert, hasta 40 s)...' -ForegroundColor DarkGray
    try { Add-Info "tracert $($Destinos[0])" (& tracert.exe -d -h 15 -w 800 $Destinos[0] 2>&1 | Out-String) } catch { }
}
[IO.File]::WriteAllText($txtInfo, $info.ToString(), $script:Utf8Bom)

if ($VelocidadCadaMinutos -gt 0) { Initialize-Http }

$muestras = New-Object System.Collections.Generic.List[object]
$pruebas = New-Object System.Collections.Generic.List[object]
$script:UltimoGateway = if ($act0) { $act0.Gateway } else { $null }
$script:UltimoDns = if ($act0) { $act0.Dns } else { $null }
$prevStats = $null
$ultimaCaptura = [datetime]::MinValue
$avisoUbicacion = $false
$proximaVelocidad = if ($VelocidadCadaMinutos -gt 0) { (Get-Date).AddMinutes(1) } else { [datetime]::MaxValue }

try {
    while ((Get-Date) -lt $fin) {
        $reloj = [Diagnostics.Stopwatch]::StartNew()
        $ahora = Get-Date
        $act = Get-InterfazActiva
        if ($act) { $script:UltimoGateway = $act.Gateway; if ($act.Dns) { $script:UltimoDns = $act.Dns } }
        $gateway = $script:UltimoGateway
        $notas = @()

        $pings = Invoke-Pings (@($gateway) + $Destinos) 1000
        $dns = Test-Dns $script:UltimoDns $NombreDns 2000

        # Tráfico del adaptador desde la muestra anterior
        $baja = $null; $sube = $null
        if ($act) {
            try {
                $st = $act.Nic.GetIPStatistics()
                if ($prevStats -and $prevStats.Id -eq $act.Nic.Id) {
                    $seg = ($ahora - $prevStats.T).TotalSeconds
                    if ($seg -gt 0) {
                        $baja = [math]::Round(($st.BytesReceived - $prevStats.Rx) * 8 / $seg / 1e6, 2)
                        $sube = [math]::Round(($st.BytesSent - $prevStats.Tx) * 8 / $seg / 1e6, 2)
                        if ($baja -lt 0 -or $sube -lt 0) { $baja = $null; $sube = $null }
                    }
                } elseif ($prevStats) {
                    $notas += 'CAMBIO_DE_ADAPTADOR'
                }
                $prevStats = @{ Id = $act.Nic.Id; Rx = $st.BytesReceived; Tx = $st.BytesSent; T = $ahora }
            } catch { $prevStats = $null }
        } else {
            $prevStats = $null
        }

        # Wi-Fi
        $wifi = $null
        if (-not $act -or $act.Tipo -eq 'Wireless80211') {
            $wifi = Get-InfoWifi
            if ($wifi.Nota) { $notas += $wifi.Nota }
            if ($wifi.Nota -eq 'SIN_PERMISO_UBICACION' -and -not $avisoUbicacion) {
                $avisoUbicacion = $true
                Write-Host 'Aviso: Windows no deja leer los datos del Wi-Fi sin permiso de ubicación.' -ForegroundColor Yellow
                Write-Host '  Actívalo en Configuración > Privacidad y seguridad > Ubicación: "Servicios de ubicación" y' -ForegroundColor Yellow
                Write-Host '  "Permitir que las aplicaciones de escritorio accedan a tu ubicación". El resto del registro sigue.' -ForegroundColor Yellow
            }
        }

        # Programas con conexiones cuando hay tráfico alto (máximo una captura por minuto)
        $procesos = ''
        if ((($null -ne $sube -and $sube -gt 2) -or ($null -ne $baja -and $baja -gt 5)) -and ($ahora - $ultimaCaptura).TotalSeconds -ge 60) {
            $procesos = Get-ProcesosConConexiones
            $ultimaCaptura = $ahora
        }

        $internetOk = @($Destinos | Where-Object { $pings[$_] -and $pings[$_].Estado -eq 'OK' }).Count -gt 0
        $pm = if ($gateway) { $pings[$gateway] } else { $null }

        $fila = [ordered]@{
            FechaHora    = $ahora.ToString('yyyy-MM-dd HH:mm:ss')
            UTC          = Get-ZonaUtc $ahora
            Conexion     = Get-NombreConexion $act
            Adaptador    = if ($act) { $act.Nic.Name } else { $null }
            Modem_IP     = $gateway
            Modem_ms     = if ($pm) { $pm.Ms } else { $null }
            Modem_estado = if ($pm) { $pm.Estado } else { 'SIN_GATEWAY' }
        }
        foreach ($d in $Destinos) {
            $fila["$($d)_ms"] = $pings[$d].Ms
            $fila["$($d)_estado"] = $pings[$d].Estado
        }
        $fila['Internet_OK'] = if ($internetOk) { 'SI' } else { 'NO' }
        $fila['DNS_servidor'] = $script:UltimoDns
        $fila['DNS_ms'] = $dns.Ms
        $fila['DNS_estado'] = $dns.Estado
        $fila['PC_baja_Mbps'] = $baja
        $fila['PC_sube_Mbps'] = $sube
        $fila['WiFi_SSID'] = if ($wifi) { $wifi.Ssid } else { $null }
        $fila['WiFi_BSSID'] = if ($wifi) { $wifi.Bssid } else { $null }
        $fila['WiFi_banda'] = if ($wifi) { $wifi.Banda } else { $null }
        $fila['WiFi_canal'] = if ($wifi) { $wifi.Canal } else { $null }
        $fila['WiFi_senal_pct'] = if ($wifi) { $wifi.Senal } else { $null }
        $fila['WiFi_rssi'] = if ($wifi) { $wifi.Rssi } else { $null }
        $fila['WiFi_radio'] = if ($wifi) { $wifi.Radio } else { $null }
        $fila['WiFi_vel_rx_Mbps'] = if ($wifi) { $wifi.Rx } else { $null }
        $fila['WiFi_vel_tx_Mbps'] = if ($wifi) { $wifi.Tx } else { $null }
        $fila['Procesos_con_conexiones'] = $procesos
        $fila['Nota'] = ($notas -join ' ')
        $obj = [pscustomobject]$fila
        $muestras.Add($obj)
        Save-Fila $csvMuestras $obj

        # Línea en pantalla
        $partes = @("Módem $(Format-Ms $pm)")
        foreach ($d in $Destinos) { $partes += "$d $(Format-Ms $pings[$d])" }
        $partes += "DNS $(Format-Ms $dns)"
        if ($wifi -and $null -ne $wifi.Senal) { $partes += ("Wi-Fi {0} c{1} {2}%" -f $wifi.Banda, $wifi.Canal, $wifi.Senal) }
        elseif ($act) { $partes += (Get-NombreConexion $act) }
        if ($null -ne $baja) { $partes += ("PC baja {0:N1} / sube {1:N1} Mbps" -f $baja, $sube) }
        $maxMs = ($Destinos | ForEach-Object { $pings[$_].Ms } | Where-Object { $null -ne $_ } | Measure-Object -Maximum).Maximum
        $algunaFalla = @($Destinos | Where-Object { $pings[$_].Estado -ne 'OK' }).Count -gt 0 -or ($pm -and $pm.Estado -ne 'OK')
        $color = if (-not $internetOk) { 'Red' } elseif ($algunaFalla -or $maxMs -gt 100) { 'Yellow' } else { 'Green' }
        Write-Host ("{0}  {1}" -f $ahora.ToString('HH:mm:ss'), ($partes -join ' | ')) -ForegroundColor $color

        # Prueba de velocidad periódica
        if ($ahora -ge $proximaVelocidad) {
            if ($internetOk) {
                $prueba = Invoke-PruebaVelocidad $act $gateway
                $pruebas.Add($prueba)
                Save-Fila $csvVelocidad $prueba
                $proximaVelocidad = (Get-Date).AddMinutes($VelocidadCadaMinutos)
                $prevStats = $null   # que el tráfico de la prueba no cuente como tráfico de la PC
                continue
            }
            $proximaVelocidad = (Get-Date).AddMinutes(1)   # sin internet: reintentar en un minuto
        }

        $restante = $IntervaloSegundos * 1000 - $reloj.ElapsedMilliseconds
        if ($restante -gt 0) { Start-Sleep -Milliseconds $restante }
    }
} finally {
    Write-Host ''
    Write-Host 'Generando resumen...' -ForegroundColor White
    $resumen = Write-Resumen
    [IO.File]::WriteAllText($txtResumen, ($resumen -join "`r`n"), $script:Utf8Bom)
    $resumen | ForEach-Object { Write-Host $_ }
    if ($script:Http) { $script:Http.Dispose() }
    $zip = "$dir.zip"
    try {
        Compress-Archive -Path (Join-Path $dir '*') -DestinationPath $zip -Force -ErrorAction Stop
        Write-Host ''
        Write-Host "Listo. Envía este archivo para analizarlo: $zip" -ForegroundColor Green
    } catch {
        Write-Host ''
        Write-Host "Listo. Los archivos están en: $dir" -ForegroundColor Green
    }
}
