# ===============================================
# ADCOS — Entra ID: Operações em CONTAS DESATIVADAS
# - RELATÓRIO preciso (FriendlyName + resumo por SkuId)
# - REMOVER licenças
# - REVOGAR sessões
# - REMOVER MFA
# - RESETAR senhas (>= 32 chars)
# ===============================================

# 0) Sessão limpa (evita overflow de função/estado)
Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue

# 1) Módulos mínimos necessários
$needed = @(
  'Microsoft.Graph.Users',                      # Get/Update-MgUser
  'Microsoft.Graph.Users.Actions',              # Get/Remove-MgUserAuthenticationMethod
  'Microsoft.Graph.Identity.SignIns',           # Revoke-MgUserSignInSession
  'Microsoft.Graph.Identity.DirectoryManagement'# Get-MgSubscribedSku / Set-MgUserLicense
)
foreach ($m in $needed) {
  if (-not (Get-Module -ListAvailable -Name $m)) {
    Write-Host ("Instalando módulo: {0}..." -f $m) -ForegroundColor Cyan
    Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
  } else {
    try { Update-Module -Name $m -ErrorAction SilentlyContinue } catch {}
  }
  Import-Module $m -ErrorAction Stop
}
Write-Host "Módulos prontos." -ForegroundColor Green

# 2) Conexão ao Graph
$scopes = @('User.Read.All','Directory.Read.All','Directory.ReadWrite.All','User.ReadWrite.All','AuditLog.Read.All')
Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes | Out-Null
$ctx = Get-MgContext
Write-Host ("Conectado ao tenant {0}" -f $ctx.TenantId) -ForegroundColor Green

# 3) Auditoria de tempo
$nowLocal  = Get-Date
$nowUTC    = $nowLocal.ToUniversalTime()
Write-Host ("Agora (LOCAL): {0} | Agora (UTC): {1}" -f $nowLocal, $nowUTC) -ForegroundColor Cyan

# 4) Buscar SOMENTE contas DESATIVADAS
$filter = "accountEnabled eq false"
$props  = 'id,displayName,userPrincipalName,accountEnabled,userType,assignedLicenses,assignedPlans'
Write-Host ("Consultando CONTAS DESATIVADAS (filtro: {0})..." -f $filter) -ForegroundColor Cyan
$usersDisabled = Get-MgUser -All -Filter $filter -Property $props -ConsistencyLevel eventual
$qtDisabled = ($usersDisabled | Measure-Object).Count
if ($qtDisabled -eq 0) {
  Write-Host "Nenhuma conta desativada encontrada. Nada a fazer." -ForegroundColor Yellow
  try { Disconnect-MgGraph -Confirm:$false | Out-Null } catch {}
  return
}
Write-Host ("Contas desativadas encontradas: {0}" -f $qtDisabled) -ForegroundColor Yellow

# 5) Dicionário de nomes amigáveis por SkuPartNumber (cobertura comum)
$friendlySku = @{
  'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'
  'O365_BUSINESS_PREMIUM'    = 'Microsoft 365 Business Standard'
  'SPB'                      = 'Microsoft 365 Business Premium'
  'M365_BUSINESS'            = 'Microsoft 365 Apps for Business'
  'STANDARDPACK'             = 'Office 365 E1'
  'ENTERPRISEPACK'           = 'Office 365 E3'
  'ENTERPRISEPREMIUM'        = 'Office 365 E5'
  'SPE_E3'                   = 'Microsoft 365 E3'
  'SPE_E5'                   = 'Microsoft 365 E5'
  'EMS'                      = 'Enterprise Mobility + Security E3'
  'EMSPREMIUM'               = 'Enterprise Mobility + Security E5'
  'POWER_BI_PRO'             = 'Power BI Pro'
  'POWER_BI_STANDARD'        = 'Power BI (Free)'
  'PROJECTPROFESSIONAL'      = 'Project Plan 3'
  'PROJECTPREMIUM'           = 'Project Plan 5'
  'VISIOCLIENT'              = 'Visio Plan 2'
  'AAD_PREMIUM'              = 'Entra ID P1'
  'AAD_PREMIUM_P2'           = 'Entra ID P2'
  'MDATP'                    = 'Defender for Endpoint'
  'MTR_PREM'                 = 'Teams Rooms Pro'
}

# 6) Mapa de SKUs do tenant: SkuId -> (SkuPartNumber, FriendlyName)
Write-Host "Mapeando SKUs do tenant..." -ForegroundColor Cyan
$skuMap = @{}  # key: SkuId (string)
try {
  $skuList = Get-MgSubscribedSku -ErrorAction Stop
} catch {
  Write-Host "Aviso: não foi possível obter SKUs; seguirei com PartNumber conhecido ou GUID." -ForegroundColor Yellow
  $skuList = @()
}
foreach ($s in $skuList) {
  if ($null -eq $s) { continue }
  $sid = $null
  try {
    if ($s.SkuId -is [guid]) { $sid = [string]$s.SkuId.Guid }
    elseif ($s.SkuId)        { $sid = [string]([guid]$s.SkuId).Guid }
  } catch { $sid = $null }
  if (-not $sid) { continue }

  $part = $null
  try { $part = [string]$s.SkuPartNumber } catch { $part = $null }
  if (-not $part) { $part = $sid }

  $fname = $friendlySku[$part]; if (-not $fname) { $fname = $part }

  $skuMap[$sid] = [PSCustomObject]@{
    SkuId         = $sid
    SkuPartNumber = $part
    FriendlyName  = $fname
  }
}
Write-Host ("SKUs mapeados: {0}" -f $skuMap.Count) -ForegroundColor Green

# 7) Pasta de saída
$outDir = 'C:\Relatorio'
if (-not (Test-Path $outDir)) {
  New-Item -Path $outDir -ItemType Directory -Force | Out-Null
  Write-Host ("Criada pasta: {0}" -f $outDir) -ForegroundColor Cyan
}

# 8) Menu
Write-Host ""
Write-Host "Escolha uma opção para CONTAS DESATIVADAS:" -ForegroundColor Cyan
Write-Host "1) Gerar RELATÓRIO das licenças (detalhado) + RESUMO por SKU (consistente)"
Write-Host "2) REMOVER TODAS as licenças das contas desativadas (gera relatório das removidas)"
Write-Host "3) Revogar sessões (sign-in sessions) de todas as contas desativadas"
Write-Host "4) Remover TODOS os MFA das contas desativadas"
Write-Host "5) Resetar a senha de todas as contas desativadas (≥ 32 chars) e forçar troca"
Write-Host "6) Sair"
$opt = Read-Host "Opção (1/2/3/4/5/6)"

# 9) Helpers inline (sem functions)
# 9.1) Extrair SKUs alinhados (ids/parts/friendly SEMPRE MESMO TAMANHO) — retorna arrays de string
$extractSkus = {
  param($assignedLicenses, $skuMap)
  $ids      = New-Object System.Collections.Generic.List[string]
  $parts    = New-Object System.Collections.Generic.List[string]
  $friendly = New-Object System.Collections.Generic.List[string]

  if ($assignedLicenses) {
    foreach ($al in $assignedLicenses) {
      $sid = $null
      try {
        if ($al.SkuId -is [guid]) { $sid = [string]$al.SkuId.Guid }
        elseif ($al.SkuId)        { $sid = [string]([guid]$al.SkuId).Guid }
      } catch { $sid = $null }
      if (-not $sid) { continue }

      $ids.Add($sid)
      if ($skuMap.ContainsKey($sid)) {
        $parts.Add(    [string]$skuMap[$sid].SkuPartNumber )
        $friendly.Add( [string]$skuMap[$sid].FriendlyName  )
      } else {
        # fallback: mantém o alinhamento (3 colunas com o mesmo comprimento)
        $parts.Add($sid)
        $friendly.Add($sid)
      }
    }
  }

  return ,($ids.ToArray()), ($parts.ToArray()), ($friendly.ToArray())
}

# 9.2) Planos de serviço (habilitados/desabilitados)
$getPlanStrings = {
  param($assignedPlans)
  if (-not $assignedPlans) { return "", "" }
  $enabled  = New-Object System.Collections.Generic.List[string]
  $disabled = New-Object System.Collections.Generic.List[string]
  foreach ($p in $assignedPlans) {
    $name = $null; $cap = $null
    try { $name = [string]$p.ServicePlanName } catch {}
    try { $cap  = [string]$p.CapabilityStatus } catch {}
    if (-not $name) { continue }
    if ($cap -and $cap -match 'Enabled') { $enabled.Add($name) } else { $disabled.Add($name) }
  }
  return ($enabled -join ';'), ($disabled -join ';')
}

# 9.3) Senha forte (>= 32)
$randPwd = {
  param([int]$length = 32)
  $upper = "ABCDEFGHJKLMNPQRSTUVWXYZ"
  $lower = "abcdefghijkmnopqrstuvwxyz"
  $digit = "23456789"
  $spec  = "!@#$%^&*()-_=+[]{}<>?"
  $all   = ($upper + $lower + $digit + $spec).ToCharArray()
  $rng   = [System.Security.Cryptography.RandomNumberGenerator]::Create()
  $bytes = New-Object byte[] ($length)
  $rng.GetBytes($bytes)
  $chars = for ($i=0; $i -lt $length; $i++) { $all[ $bytes[$i] % $all.Length ] }
  $chars[0] = $upper[ [int]($bytes[0] % $upper.Length) ]
  $chars[1] = $lower[ [int]($bytes[1] % $lower.Length) ]
  $chars[2] = $digit[ [int]($bytes[2] % $digit.Length) ]
  $chars[3] = $spec[  [int]($bytes[3] % $spec.Length) ]
  -join ($chars | Sort-Object {Get-Random})
}

# 10) Opções
if ($opt -eq '1') {
  # ===== RELATÓRIO DETALHADO + RESUMO (consistentes) =====
  $rows = New-Object System.Collections.Generic.List[object]
  $aggBySkuId = @{}  # key: SkuId (string) -> int count

  foreach ($u in $usersDisabled) {
    $ids,$parts,$friendly = & $extractSkus $u.AssignedLicenses $skuMap
    $enPlans,$disPlans    = & $getPlanStrings $u.AssignedPlans

    # Agregação por SkuId — sem '++' (soma explícita com int)
    foreach ($sid in $ids) {
      $key = [string]$sid
      if (-not $aggBySkuId.ContainsKey($key)) { $aggBySkuId[$key] = 0 }
      $aggBySkuId[$key] = [int]$aggBySkuId[$key] + 1
    }

    $rows.Add([PSCustomObject]@{
      SnapshotLocalTime    = $nowLocal
      SnapshotUTC          = $nowUTC
      DisplayName          = $u.DisplayName
      UserPrincipalName    = $u.UserPrincipalName
      Id                   = $u.Id
      AccountEnabled       = $u.AccountEnabled
      UserType             = $u.UserType
      LicenseCount         = $ids.Count
      LicenseSkuIds        = ($ids -join ';')
      LicenseSkuPartNums   = ($parts -join ';')
      LicenseFriendlyNames = ($friendly -join ';')
      PlansEnabled         = $enPlans
      PlansDisabled        = $disPlans
    })
  }

  $rows = $rows | Sort-Object -Property @{Expression='LicenseCount';Descending=$true}, UserPrincipalName

  $csvDet = Join-Path $outDir ("Relatorio_Licencas_ContasDesativadas_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
  $rows | Export-Csv -Path $csvDet -NoTypeInformation -Encoding UTF8
  Write-Host ("Relatório detalhado salvo: {0}" -f $csvDet) -ForegroundColor Green

  # RESUMO por SkuId com enriquecimento
  $sum = New-Object System.Collections.Generic.List[object]
  foreach ($sid in $aggBySkuId.Keys) {
    $part = $sid; $fname = $sid
    if ($skuMap.ContainsKey($sid)) {
      $part  = $skuMap[$sid].SkuPartNumber
      $fname = $skuMap[$sid].FriendlyName
    }
    $sum.Add([PSCustomObject]@{
      SkuId          = $sid
      SkuPartNumber  = $part
      FriendlyName   = $fname
      DisabledUsers  = [int]$aggBySkuId[$sid]
    })
  }
  $csvSum = Join-Path $outDir ("Resumo_Licencas_ContasDesativadas_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
  $sum | Sort-Object -Property DisabledUsers -Descending | Export-Csv -Path $csvSum -NoTypeInformation -Encoding UTF8
  Write-Host ("Resumo por SKU salvo: {0}" -f $csvSum) -ForegroundColor Cyan

  # SANITY CHECK de consistência
  $totalAssignments = ($rows | Measure-Object -Property LicenseCount -Sum).Sum
  $sumAssignments   = ($sum  | Measure-Object -Property DisabledUsers -Sum).Sum
  if ($totalAssignments -eq $sumAssignments) {
    Write-Host ("[OK] Consistência: detalhado LicenseCount sum = resumo DisabledUsers sum = {0}" -f $totalAssignments) -ForegroundColor Green
  } else {
    Write-Host ("[ALERTA] Inconsistência: detalhado = {0} | resumo = {1}" -f $totalAssignments, $sumAssignments) -ForegroundColor Red
  }
}
elseif ($opt -eq '2') {
  # ===== REMOVER TODAS AS LICENÇAS =====
  Write-Host "ATENÇÃO: isto removerá TODAS as licenças de todas as contas DESATIVADAS." -ForegroundColor Red
  $confirm = Read-Host "Digite 'S' para confirmar (ou 'N' p/ cancelar)"
  if ($confirm -match '^[sS]$') {
    $log = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($u in $usersDisabled) {
      $i++
      $ids,$parts,$friendly = & $extractSkus $u.AssignedLicenses $skuMap
      Write-Progress -Activity "Removendo licenças" -Status ("{0}/{1}: {2}" -f $i,$qtDisabled,$u.UserPrincipalName) -PercentComplete (($i/$qtDisabled)*100)

      if ($ids.Count -eq 0) {
        $log.Add([PSCustomObject]@{
          Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName;
          RemovedSkuIds=''; RemovedSkuPartNumbers=''; RemovedFriendlyNames=''; Result='Sem licenças'
        })
        continue
      }

      try {
        Set-MgUserLicense -UserId $u.Id -AddLicenses @() -RemoveLicenses $ids -ErrorAction Stop
        $log.Add([PSCustomObject]@{
          Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName;
          RemovedSkuIds=($ids -join ';'); RemovedSkuPartNumbers=($parts -join ';'); RemovedFriendlyNames=($friendly -join ';'); Result='Removidas'
        })
        Write-Host ("Licenças removidas de: {0}" -f $u.UserPrincipalName) -ForegroundColor Green
      } catch {
        $msg = $_.Exception.Message
        $log.Add([PSCustomObject]@{
          Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName;
          RemovedSkuIds=($ids -join ';'); RemovedSkuPartNumbers=($parts -join ';'); RemovedFriendlyNames=($friendly -join ';'); Result=("Erro: {0}" -f $msg)
        })
        Write-Host ("Erro ao remover licenças de {0}: {1}" -f $u.UserPrincipalName, $msg) -ForegroundColor Red
      }
    }
    $csv = Join-Path $outDir ("Remocao_Licencas_ContasDesativadas_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
    $log | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Host ("Relatório salvo: {0}" -f $csv) -ForegroundColor Cyan
  } else {
    Write-Host "Operação cancelada." -ForegroundColor Yellow
  }
}
elseif ($opt -eq '3') {
  # ===== REVOGAR SESSÕES =====
  Write-Host "Revogando sessões de TODAS as contas DESATIVADAS..." -ForegroundColor Yellow
  $log = New-Object System.Collections.Generic.List[object]
  $i = 0
  foreach ($u in $usersDisabled) {
    $i++
    Write-Progress -Activity "Revogando sessões" -Status ("{0}/{1}: {2}" -f $i,$qtDisabled,$u.UserPrincipalName) -PercentComplete (($i/$qtDisabled)*100)
    try {
      Revoke-MgUserSignInSession -UserId $u.Id -ErrorAction Stop
      $log.Add([PSCustomObject]@{ Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName; Result='Sessões revogadas' })
      Write-Host ("Sessões revogadas: {0}" -f $u.UserPrincipalName) -ForegroundColor Green
    } catch {
      $msg = $_.Exception.Message
      $log.Add([PSCustomObject]@{ Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName; Result=("Erro: {0}" -f $msg) })
      Write-Host ("Erro ao revogar sessões de {0}: {1}" -f $u.UserPrincipalName, $msg) -ForegroundColor Red
    }
  }
  $csv = Join-Path $outDir ("Revogar_Sessoes_ContasDesativadas_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
  $log | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
  Write-Host ("Relatório salvo: {0}" -f $csv) -ForegroundColor Cyan
}
elseif ($opt -eq '4') {
  # ===== REMOVER MFA =====
  Write-Host "ATENÇÃO: isto removerá TODOS os métodos de MFA das contas DESATIVADAS." -ForegroundColor Red
  $confirm = Read-Host "Digite 'S' para confirmar (ou 'N' p/ cancelar)"
  if ($confirm -match '^[sS]$') {
    $log = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($u in $usersDisabled) {
      $i++
      Write-Progress -Activity "Removendo MFA" -Status ("{0}/{1}: {2}" -f $i,$qtDisabled,$u.UserPrincipalName) -PercentComplete (($i/$qtDisabled)*100)
      try {
        $methods = Get-MgUserAuthenticationMethod -UserId $u.Id -ErrorAction Stop
        if (-not $methods -or $methods.Count -eq 0) {
          $log.Add([PSCustomObject]@{ Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName; MethodId=''; MethodType=''; Result='Sem métodos' })
          continue
        }
        foreach ($m in $methods) {
          $type = $m.AdditionalProperties['@odata.type']
          $mid  = $m.Id
          try {
            Remove-MgUserAuthenticationMethod -UserId $u.Id -AuthenticationMethodId $mid -ErrorAction Stop
            $log.Add([PSCustomObject]@{ Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName; MethodId=$mid; MethodType=$type; Result='Removido' })
          } catch {
            $msg2 = $_.Exception.Message
            $log.Add([PSCustomObject]@{ Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName; MethodId=$mid; MethodType=$type; Result=("Erro: {0}" -f $msg2) })
          }
        }
        Write-Host ("MFA removido (onde existia) para: {0}" -f $u.UserPrincipalName) -ForegroundColor Green
      } catch {
        $msg = $_.Exception.Message
        $log.Add([PSCustomObject]@{ Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName; MethodId=''; MethodType=''; Result=("Erro geral: {0}" -f $msg) })
        Write-Host ("Erro ao listar/remover MFA de {0}: {1}" -f $u.UserPrincipalName, $msg) -ForegroundColor Red
      }
    }
    $csv = Join-Path $outDir ("Remocao_MFA_ContasDesativadas_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
    $log | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Host ("Relatório salvo: {0}" -f $csv) -ForegroundColor Cyan
  } else {
    Write-Host "Operação cancelada." -ForegroundColor Yellow
  }
}
elseif ($opt -eq '5') {
  # ===== RESETAR SENHAS =====
  Write-Host "ATENÇÃO: isto irá RESETAR a senha de TODAS as contas DESATIVADAS." -ForegroundColor Red
  $confirm = Read-Host "Digite 'S' para confirmar (ou 'N' p/ cancelar)"
  if ($confirm -match '^[sS]$') {
    $log = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($u in $usersDisabled) {
      $i++
      Write-Progress -Activity "Resetando senhas" -Status ("{0}/{1}: {2}" -f $i,$qtDisabled,$u.UserPrincipalName) -PercentComplete (($i/$qtDisabled)*100)
      try {
        $pwd = & $randPwd 32
        $profile = @{ password = $pwd; forceChangePasswordNextSignIn = $true }
        Update-MgUser -UserId $u.Id -PasswordProfile $profile -ErrorAction Stop
        $log.Add([PSCustomObject]@{ Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName; NewPassword=$pwd; Result='Senha resetada' })
        Write-Host ("Senha resetada: {0}" -f $u.UserPrincipalName) -ForegroundColor Green
      } catch {
        $msg = $_.Exception.Message
        $log.Add([PSCustomObject]@{ Time=(Get-Date).ToString('u'); UserPrincipalName=$u.UserPrincipalName; NewPassword=''; Result=("Erro: {0}" -f $msg) })
        Write-Host ("Erro ao resetar senha de {0}: {1}" -f $u.UserPrincipalName, $msg) -ForegroundColor Red
      }
    }
    $csv = Join-Path $outDir ("Reset_Senhas_ContasDesativadas_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv")
    $log | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Host ("Relatório salvo: {0}" -f $csv) -ForegroundColor Cyan
    Write-Host "ATENÇÃO: proteja este arquivo, pois contém senhas em texto claro." -ForegroundColor Red
  } else {
    Write-Host "Operação cancelada." -ForegroundColor Yellow
  }
}
else {
  Write-Host "Saindo..." -ForegroundColor Cyan
}

# 11) Desconectar
try { Disconnect-MgGraph -Confirm:$false | Out-Null } catch {}
Write-Host "Concluído." -ForegroundColor Green
