# =======================
# Entra ID — Contas ATIVAS sem logon há X dias (preciso, estável, auditável)
# =======================

# 0) Limpeza de sessão
Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue

# 1) Importar SOMENTE submódulos necessários
$needed = @('Microsoft.Graph.Users','Microsoft.Graph.Identity.SignIns')
foreach ($m in $needed) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host ("Instalando módulo: {0}..." -f $m) -ForegroundColor Cyan
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    } else {
        try { Update-Module -Name $m -ErrorAction SilentlyContinue } catch { }
    }
    Import-Module $m -ErrorAction Stop
}
Write-Host "Módulos prontos." -ForegroundColor Green

# 2) Conectar ao Graph
$scopes = @('User.Read.All','Directory.Read.All','AuditLog.Read.All','User.ReadWrite.All','Directory.ReadWrite.All')
Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes | Out-Null
$ctx = Get-MgContext
Write-Host ("Conectado ao tenant {0}" -f $ctx.TenantId) -ForegroundColor Green

# 3) Parâmetro de dias
$dias = $null
while (-not $dias) {
    $in = Read-Host "Informe o período em dias (ex.: 45) — usuários ATIVOS que NÃO autenticam há X dias"
    if ([int]::TryParse($in, [ref]([int]$null))) { $dias = [int]$in }
    if (-not $dias -or $dias -le 0) { Write-Host "Valor inválido. Digite um inteiro > 0." -ForegroundColor Red; $dias = $null }
}

# 4) Hora local e cutoff (auditáveis)
$nowLocal  = Get-Date
$nowUTC    = $nowLocal.ToUniversalTime()
$cutoffUTC = $nowUTC.AddDays(-$dias)
Write-Host ("Agora (LOCAL): {0} | Agora (UTC): {1} | Cutoff (UTC): {2}" -f $nowLocal, $nowUTC, $cutoffUTC) -ForegroundColor Cyan

# 5) Trazer apenas contas ATIVAS
$filter = "accountEnabled eq true"
$props  = 'id,displayName,userPrincipalName,accountEnabled,userType,createdDateTime,signInActivity'
Write-Host ("Consultando usuários (filtro: {0})..." -f $filter) -ForegroundColor Cyan
$users = Get-MgUser -All -Filter $filter -Property $props -ConsistencyLevel eventual

# 6) Parser robusto p/ UTC
$style = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
$cult  = [System.Globalization.CultureInfo]::InvariantCulture
$toUtc = {
    param($s, $cult, $style)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $dto = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse(($s -as [string]), $cult, $style, [ref]$dto)) {
        $dt = $dto.UtcDateTime
        if ($dt.Year -lt 2000) { return $null } else { return $dt }
    }
    return $null
}

# 7) Consolidar logons e aplicar REGRA (≥ X dias)
$rows = @()
foreach ($u in $users) {
    $succ = $null; $inter = $null; $noni = $null; $created = $null
    $neverSignedIn = $false

    if ($u.SignInActivity) {
        $succ  = & $toUtc $u.SignInActivity.LastSuccessfulSignInDateTime     $cult $style
        $inter = & $toUtc $u.SignInActivity.LastInteractiveSignInDateTime     $cult $style
        $noni  = & $toUtc $u.SignInActivity.LastNonInteractiveSignInDateTime  $cult $style
    }
    if ($u.PSObject.Properties.Name -contains 'CreatedDateTime' -and $u.CreatedDateTime) {
        $created = & $toUtc $u.CreatedDateTime $cult $style
    }

    # último logon real (máximo) + origem
    $cands = @()
    if ($succ)  { $cands += [PSCustomObject]@{Name='Successful';     When=$succ} }
    if ($inter) { $cands += [PSCustomObject]@{Name='Interactive';    When=$inter} }
    if ($noni)  { $cands += [PSCustomObject]@{Name='NonInteractive'; When=$noni} }

    $lastSeen = $null; $lastSource = $null
    if ($cands.Count -gt 0) {
        $max = $cands | Sort-Object When -Descending | Select-Object -First 1
        $lastSeen = $max.When; $lastSource = $max.Name
    }

    # Precisão: "≥ X dias"  == lastSeen <= cutoff
    $isInactive = $false
    $status = 'Sem dados suficientes'

    if ($lastSeen) {
        if ($lastSeen -le $cutoffUTC) {
            $isInactive = $true
            $status = "Inativo (≥ $dias dias sem logon)"
        } else {
            $isInactive = $false
            $status = "Ativo (último: $($lastSeen.ToString('u')) via $lastSource)"
        }
    } else {
        if ($created -and $created -le $cutoffUTC) {
            # >>>> MENSAGEM AMIGÁVEL (pedido) <<<<
            $neverSignedIn = $true
            $isInactive = $true
            $daysSinceCreated = [math]::Round(($nowUTC - $created).TotalDays, 1)
            $status = ("Sem logon registrado desde a criação (criada em {0:u}, há {1} dias)" -f $created, $daysSinceCreated)
        } else {
            $isInactive = $false
            if ($created) { $status = "Sem logon ainda (criada em $($created.ToString('u')))" }
        }
    }

    $daysSince = $null
    if ($lastSeen) { $daysSince = [math]::Round(($nowUTC - $lastSeen).TotalDays, 1) }

    $rows += [PSCustomObject]@{
        AdminInputDays              = $dias
        NowLocal                    = $nowLocal
        NowUTC                      = $nowUTC
        CutoffUTC                   = $cutoffUTC
        DisplayName                 = $u.DisplayName
        UserPrincipalName           = $u.UserPrincipalName
        Id                          = $u.Id
        UserType                    = $u.UserType
        AccountEnabled              = $u.AccountEnabled
        CreatedUTC                  = $created
        LastSuccessfulSignInUTC     = $succ
        LastInteractiveSignInUTC    = $inter
        LastNonInteractiveSignInUTC = $noni
        LastSeenUTC                 = $lastSeen
        LastSeenSource              = $lastSource
        DaysSinceLastSeen           = $daysSince
        NeverSignedIn               = $neverSignedIn
        IsInactive                  = $isInactive
        Status                      = $status
    }
}

# 8) Somente INATIVOS (precisos)
$inativos = $rows | Where-Object { $_.IsInactive -eq $true } | Sort-Object UserPrincipalName
$qtInativos = ($inativos | Measure-Object).Count
Write-Host ("Inativos (precisos) encontrados: {0}" -f $qtInativos) -ForegroundColor Yellow

# 9) Menu
Write-Host ""
Write-Host "Opções:" -ForegroundColor Cyan
Write-Host "1) Gerar RELATÓRIO (somente INATIVOS) em C:\Relatorio"
Write-Host "2) Bloquear todas as contas INATIVAS listadas (AccountEnabled=false)"
Write-Host "3) Sair e desconectar"
$opt = Read-Host "Opção (1/2/3)"

# 10) Relatório (somente inativos) + completo p/ auditoria
if ($opt -eq '1') {
    $outDir = 'C:\Relatorio'
    if (-not (Test-Path $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }

    $csv = Join-Path $outDir ("Relatorio_Inativos_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
    $inativos | Select-Object AdminInputDays,NowLocal,NowUTC,CutoffUTC,
                               DisplayName,UserPrincipalName,UserType,CreatedUTC,
                               LastSuccessfulSignInUTC,LastInteractiveSignInUTC,LastNonInteractiveSignInUTC,
                               LastSeenUTC,LastSeenSource,DaysSinceLastSeen,NeverSignedIn,Status |
        Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Host ("Relatório salvo: {0}" -f $csv) -ForegroundColor Green

    $csvFull = Join-Path $outDir ("Relatorio_Completo_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
    $rows | Export-Csv -Path $csvFull -NoTypeInformation -Encoding UTF8
    Write-Host ("Relatório completo (auditoria) salvo: {0}" -f $csvFull) -ForegroundColor Cyan
}
elseif ($opt -eq '2') {
    if ($qtInativos -eq 0) {
        Write-Host "Nada a bloquear (nenhum inativo)." -ForegroundColor Yellow
    } else {
        Write-Host ("ATENÇÃO: Isto desativará {0} conta(s)." -f $qtInativos) -ForegroundColor Red
        $confirm = Read-Host "Digite 'S' para confirmar (ou 'N' p/ cancelar)"
        if ($confirm -match '^[sS]$') {
            $outDir = 'C:\Relatorio'
            if (-not (Test-Path $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }
            $fileLog = Join-Path $outDir ("Log_Bloqueio_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
            $log = New-Object System.Collections.Generic.List[object]

            $i = 0
            foreach ($u in $inativos) {
                $i++
                $target = $u.UserPrincipalName
                Write-Progress -Activity "Bloqueando contas..." -Status ("{0}/{1}: {2}" -f $i,$qtInativos,$target) -PercentComplete (($i/$qtInativos)*100)
                try {
                    Update-MgUser -UserId $u.Id -AccountEnabled:$false -ErrorAction Stop
                    $log.Add([PSCustomObject]@{ UserPrincipalName=$target; Result='Bloqueado'; Time=(Get-Date).ToString('u') })
                    Write-Host ("Bloqueado: {0}" -f $target) -ForegroundColor Green
                } catch {
                    $msg = $_.Exception.Message
                    $log.Add([PSCustomObject]@{ UserPrincipalName=$target; Result=("Erro: {0}" -f $msg); Time=(Get-Date).ToString('u') })
                    Write-Host ("Erro ao bloquear {0}: {1}" -f $target, $msg) -ForegroundColor Red
                }
            }

            $log | Export-Csv -Path $fileLog -NoTypeInformation -Encoding UTF8
            Write-Host ("Log salvo: {0}" -f $fileLog) -ForegroundColor Cyan
        } else {
            Write-Host "Operação cancelada." -ForegroundColor Yellow
        }
    }
}

# 11) Desconectar
try { Disconnect-MgGraph -Confirm:$false | Out-Null } catch {}
Write-Host "Concluído." -ForegroundColor Green
