# ==========================================================
# Entra ID — Inativos ≥ X dias — v4 (rápido: /users + $batch para audit logs)
# ==========================================================
$ErrorActionPreference = 'Stop'
Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue

# 1) Módulos
$needed = @('Microsoft.Graph.Authentication','Microsoft.Graph.Users','Microsoft.Graph.Identity.SignIns')
foreach ($m in $needed) {
  if (-not (Get-Module -ListAvailable -Name $m)) {
    Write-Host ("Instalando módulo: {0}..." -f $m) -ForegroundColor Cyan
    Install-Module -Name $m -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
  }
  Import-Module $m -ErrorAction Stop
}
Write-Host "Módulos prontos." -ForegroundColor Green

# 2) Conectar leitura
$scopesRead = @('User.Read.All','Directory.Read.All','AuditLog.Read.All')
Write-Host "Conectando ao Microsoft Graph (somente leitura)..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopesRead | Out-Null
$ctx = Get-MgContext
Write-Host ("Conectado ao tenant {0}" -f $ctx.TenantId) -ForegroundColor Green

# 3) Dias
$dias = $null
while (-not $dias) {
  $in = Read-Host "Informe o período em dias (ex.: 45) — usuários ATIVOS que NÃO autenticam há X dias"
  if ($in -match '^\d+$' -and [int]$in -gt 0) { $dias = [int]$in }
  if (-not $dias -or $dias -le 0) { Write-Host "Valor inválido. Digite um inteiro > 0." -ForegroundColor Red; $dias = $null }
}

# 4) Datas
$nowLocal = Get-Date
$nowUTC   = $nowLocal.ToUniversalTime()
$cutoffUTC = $nowUTC.AddDays(-$dias)
Write-Host ("Agora (LOCAL): {0} | Agora (UTC): {1} | Cutoff (UTC): {2}" -f $nowLocal,$nowUTC,$cutoffUTC) -ForegroundColor Cyan

# ===== Helpers =====
function Invoke-GraphRest {
  param([string]$Uri,[string]$Method='GET',[object]$Body=$null,[int]$MaxRetries=5,[int]$InitialDelayMs=400)
  $delay=$InitialDelayMs
  for($i=0;$i -lt $MaxRetries;$i++){
    try{
      if($Method -eq 'GET'){ return Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputType PSObject -ErrorAction Stop }
      else { return Invoke-MgGraphRequest -Uri $Uri -Method POST -Body $Body -OutputType PSObject -ErrorAction Stop }
    }catch{
      $status = $_.Exception.InnerException.ResponseStatusCode
      if($status -in 429,503){ Start-Sleep -Milliseconds $delay; $delay=[Math]::Min($delay*2,6000)+(Get-Random -Min 0 -Max 200); continue }
      throw
    }
  }
  throw "Falha após $MaxRetries tentativas: $Method $Uri"
}

function Get-GraphPaged {
  param([string]$Uri)
  $all = New-Object System.Collections.Generic.List[object]
  $page = Invoke-GraphRest -Uri $Uri
  if($page.value){ $all.AddRange($page.value) }
  $count=$all.Count; Write-Host ("  > Coletados {0} usuários..." -f $count) -ForegroundColor DarkCyan
  while($page.'@odata.nextLink'){
    $page = Invoke-GraphRest -Uri $page.'@odata.nextLink'
    if($page.value){ $all.AddRange($page.value); $count=$all.Count; Write-Host ("  > Coletados {0} usuários..." -f $count) -ForegroundColor DarkCyan }
  }
  return ,$all
}

$style=[System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
$cult =[System.Globalization.CultureInfo]::InvariantCulture
$toUtc = {
  param($s,$cult,$style)
  if([string]::IsNullOrWhiteSpace($s)){return $null}
  $dto=[datetimeoffset]::MinValue
  if([datetimeoffset]::TryParse(($s -as [string]),$cult,$style,[ref]$dto)){
    $dt=$dto.UtcDateTime; if($dt.Year -lt 2000){return $null}else{return $dt}
  } ; return $null
}

# --------- $batch para último logon em auditLogs/signIns ---------
function Get-LastSignInBatch {
  param([object[]]$Users,[int]$BatchSize=20)
  # Retorna hashtable: userId -> @{When=..., Source=...}
  $map = @{}
  $chunks = [System.Linq.Enumerable]::ToList([System.Linq.Enumerable]::Chunk($Users,$BatchSize))
  $totalBatches = $chunks.Count; $b=0
  foreach($chunk in $chunks){
    $b++; Write-Host ("  > Consultando audit logs (lote {0}/{1}, {2} usuários)..." -f $b,$totalBatches,$chunk.Count) -ForegroundColor DarkYellow
    $requests = @()
    $i=0
    foreach($u in $chunk){
      $i++
      $id = "$($u.id)"
      $flt = "userId eq '$id'"
      $uri = "/auditLogs/signIns?`$filter=$([uri]::EscapeDataString($flt))&`$orderby=createdDateTime desc&`$top=1"
      $requests += @{
        id   = "$i"
        method = "GET"
        url  = $uri
      }
    }
    $body = @{ requests = $requests } | ConvertTo-Json -Depth 6
    $resp = Invoke-GraphRest -Uri "https://graph.microsoft.com/v1.0/`$batch" -Method POST -Body $body
    foreach($r in $resp.responses){
      $idx = [int]$r.id - 1
      $u = $chunk[$idx]
      if($r.status -eq 200 -and $r.body.value.Count -gt 0){
        $rec = $r.body.value[0]
        $when = & $toUtc $rec.createdDateTime $cult $style
        $src  = if ($rec.isInteractive) { 'Interactive' } else { 'NonInteractive' }
        $map[$u.id] = @{ When=$when; Source=$src }
      } else {
        # pode ser sem eventos de logon
        $map[$u.id] = $null
      }
    }
  }
  return $map
}

# 6) Buscar usuários ATIVOS
$filter = "accountEnabled eq true"
$selectWith    = "id,displayName,userPrincipalName,accountEnabled,userType,createdDateTime,signInActivity"
$selectWithout = "id,displayName,userPrincipalName,accountEnabled,userType,createdDateTime"
$top=200

Write-Host ("Consultando usuários (filtro: {0})..." -f $filter) -ForegroundColor Cyan
$users = @(); $hasSignInActivity = $true
try{
  $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$([uri]::EscapeDataString($filter))&`$select=$([uri]::EscapeDataString($selectWith))&`$top=$top"
  $users = Get-GraphPaged -Uri $uri
}catch{
  $hasSignInActivity=$false
  Write-Host "signInActivity indisponível via /users. Usarei audit logs com $batch (rápido)." -ForegroundColor Yellow
  $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$([uri]::EscapeDataString($filter))&`$select=$([uri]::EscapeDataString($selectWithout))&`$top=$top"
  $users = Get-GraphPaged -Uri $uri
}

# 7) Se não temos signInActivity, resolva último logon com $batch
$lastSignInMap = @{}
if (-not $hasSignInActivity) {
  $lastSignInMap = Get-LastSignInBatch -Users $users -BatchSize 20
}

# 8) Consolidação
$rows = @()
foreach($u in $users){
  $succ=$null;$inter=$null;$noni=$null;$created=$null;$neverSignedIn=$false
  if($u.createdDateTime){ $created = & $toUtc $u.createdDateTime $cult $style }

  if ($hasSignInActivity -and $u.signInActivity){
    $succ  = & $toUtc $u.signInActivity.lastSuccessfulSignInDateTime     $cult $style
    $inter = & $toUtc $u.signInActivity.lastInteractiveSignInDateTime     $cult $style
    $noni  = & $toUtc $u.signInActivity.lastNonInteractiveSignInDateTime  $cult $style
  } else {
    $ls = $lastSignInMap[$u.id]
    if ($ls -and $ls.When){
      if($ls.Source -eq 'Interactive'){ $inter = $ls.When } else { $noni = $ls.When }
      $succ = $ls.When
    }
  }

  $cands=@(); if($succ){$cands+=@{Name='Successful';When=$succ}}
  if($inter){$cands+=@{Name='Interactive';When=$inter}}
  if($noni){$cands+=@{Name='NonInteractive';When=$noni}}

  $lastSeen=$null;$lastSource=$null
  if($cands.Count -gt 0){
    $max = $cands | Sort-Object When -Descending | Select-Object -First 1
    $lastSeen=$max.When;$lastSource=$max.Name
  }

  $isInactive=$false;$status='Sem dados suficientes'
  if($lastSeen){
    if($lastSeen -le $cutoffUTC){ $isInactive=$true; $status="Inativo (≥ $dias dias sem logon)" }
    else { $status = "Ativo (último: $($lastSeen.ToString('u')) via $lastSource)" }
  } else {
    if($created -and $created -le $cutoffUTC){
      $neverSignedIn=$true; $isInactive=$true
      $daysSinceCreated=[math]::Round(($nowUTC - $created).TotalDays,1)
      $status=("Sem logon registrado desde a criação (criada em {0:u}, há {1} dias)" -f $created,$daysSinceCreated)
    } else {
      if($created){ $status="Sem logon ainda (criada em $($created.ToString('u')))" }
    }
  }

  $daysSince=$null; if($lastSeen){ $daysSince=[math]::Round(($nowUTC - $lastSeen).TotalDays,1) }

  $rows += [PSCustomObject]@{
    AdminInputDays=$dias; NowLocal=$nowLocal; NowUTC=$nowUTC; CutoffUTC=$cutoffUTC
    DisplayName=$u.displayName; UserPrincipalName=$u.userPrincipalName; Id=$u.id
    UserType=$u.userType; AccountEnabled=$u.accountEnabled; CreatedUTC=$created
    LastSuccessfulSignInUTC=$succ; LastInteractiveSignInUTC=$inter; LastNonInteractiveSignInUTC=$noni
    LastSeenUTC=$lastSeen; LastSeenSource=$lastSource; DaysSinceLastSeen=$daysSince
    NeverSignedIn=$neverSignedIn; IsInactive=$isInactive; Status=$status
  }
}

# 9) Somente inativos
$inativos = $rows | Where-Object { $_.IsInactive -eq $true } | Sort-Object UserPrincipalName
$qtInativos = ($inativos | Measure-Object).Count
Write-Host ("Inativos (precisos) encontrados: {0}" -f $qtInativos) -ForegroundColor Yellow

# 10) Menu
Write-Host ""
Write-Host "Opções:" -ForegroundColor Cyan
Write-Host "1) Gerar RELATÓRIO (somente INATIVOS) em C:\Relatorio"
Write-Host "2) Bloquear todas as contas INATIVAS listadas (AccountEnabled=false)"
Write-Host "3) Sair e desconectar"
$opt = Read-Host "Opção (1/2/3)"

# 11) Relatórios / bloqueio
if ($opt -eq '1') {
  $outDir='C:\Relatorio'; if(-not (Test-Path $outDir)){ New-Item -Path $outDir -ItemType Directory -Force | Out-Null }
  $csv = Join-Path $outDir ("Relatorio_Inativos_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
  $inativos | Select-Object AdminInputDays,NowLocal,NowUTC,CutoffUTC,DisplayName,UserPrincipalName,UserType,CreatedUTC,
                             LastSuccessfulSignInUTC,LastInteractiveSignInUTC,LastNonInteractiveSignInUTC,
                             LastSeenUTC,LastSeenSource,DaysSinceLastSeen,NeverSignedIn,Status |
    Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
  Write-Host ("Relatório salvo: {0}" -f $csv) -ForegroundColor Green

  $csvFull = Join-Path $outDir ("Relatorio_Completo_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
  $rows | Export-Csv -Path $csvFull -NoTypeInformation -Encoding UTF8
  Write-Host ("Relatório completo (auditoria) salvo: {0}" -f $csvFull) -ForegroundColor Cyan
}
elseif ($opt -eq '2') {
  if ($qtInativos -eq 0) { Write-Host "Nada a bloquear (nenhum inativo)." -ForegroundColor Yellow }
  else {
    Write-Host ("ATENÇÃO: Isto desativará {0} conta(s)." -f $qtInativos) -ForegroundColor Red
    $confirm = Read-Host "Digite 'S' para confirmar (ou 'N' p/ cancelar)"
    if ($confirm -match '^[sS]$') {
      $scopesWrite=@('User.ReadWrite.All','Directory.ReadWrite.All')
      $have=(Get-MgContext).Scopes
      if (@($scopesWrite | Where-Object { $_ -notin $have }).Count -gt 0) {
        Write-Host "Solicitando escopos de escrita..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes ($scopesRead + $scopesWrite) | Out-Null
      }
      $outDir='C:\Relatorio'; if(-not (Test-Path $outDir)){ New-Item -Path $outDir -ItemType Directory -Force | Out-Null }
      $fileLog = Join-Path $outDir ("Log_Bloqueio_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
      $log = New-Object System.Collections.Generic.List[object]

      $i=0
      foreach($u in $inativos){
        $i++; $target=$u.UserPrincipalName
        Write-Progress -Activity "Bloqueando contas..." -Status ("{0}/{1}: {2}" -f $i,$qtInativos,$target) -PercentComplete (($i/$qtInativos)*100)
        try{
          Update-MgUser -UserId $u.Id -AccountEnabled:$false -ErrorAction Stop
          $log.Add([PSCustomObject]@{ UserPrincipalName=$target; Result='Bloqueado'; Time=(Get-Date).ToString('u') })
          Write-Host ("Bloqueado: {0}" -f $target) -ForegroundColor Green
        }catch{
          $msg=$_.Exception.Message
          $log.Add([PSCustomObject]@{ UserPrincipalName=$target; Result=("Erro: {0}" -f $msg); Time=(Get-Date).ToString('u') })
          Write-Host ("Erro ao bloquear {0}: {1}" -f $target,$msg) -ForegroundColor Red
        }
      }
      $log | Export-Csv -Path $fileLog -NoTypeInformation -Encoding UTF8
      Write-Host ("Log salvo: {0}" -f $fileLog) -ForegroundColor Cyan
    } else {
      Write-Host "Operação cancelada." -ForegroundColor Yellow
    }
  }
}

# 12) Desconectar (sem -Confirm)
try { Disconnect-MgGraph | Out-Null } catch {}
Write-Host "Concluído." -ForegroundColor Green
