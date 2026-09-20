# Add GitHub collaborators with Write (push) permission.
# Usage: .\add_collaborators.ps1 -Token <PAT> [-Permission push]
param(
  [Parameter(Mandatory = $true)][string]$Token,
  [string]$Permission = 'push',
  [string]$Owner = 'rrafiya',
  [string]$Repo = 'git_course'
)

$ErrorActionPreference = 'Continue'
$proxy = 'http://127.0.0.1:14138'
$users = @('sunzhi-01', 'liulangaaa', 'lhm-11')

$headers = @{
  'User-Agent'    = 'dsh'
  'Authorization' = "Bearer $Token"
  'Accept'        = 'application/vnd.github+json'
}

function Show-Error($ex) {
  if ($ex.Response) {
    $code = [int]$ex.Response.StatusCode
    $body = ''
    try {
      $sr = New-Object System.IO.StreamReader($ex.Response.GetResponseStream())
      $body = $sr.ReadToEnd()
    } catch { }
    Write-Output ("      HTTP " + $code + "  " + $body)
  } else {
    Write-Output ("      " + $ex.Message)
  }
}

# ---- 1. verify token ----
Write-Output "=== 1. verify token ==="
try {
  $me = Invoke-RestMethod -Uri 'https://api.github.com/user' -Headers $headers -Proxy $proxy -TimeoutSec 25
  Write-Output ("  OK  authenticated as: " + $me.login)
  if ($me.login -ne $Owner) {
    Write-Output ("  WARNING: token belongs to '" + $me.login + "', not '" + $Owner + "'")
  }
} catch {
  Write-Output "  FAILED - token invalid or expired"
  Show-Error $_.Exception
  exit 1
}

# ---- 2. verify repo admin right ----
Write-Output ""
Write-Output "=== 2. check repo permissions ==="
try {
  $r = Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo" -Headers $headers -Proxy $proxy -TimeoutSec 25
  Write-Output ("  repo: " + $r.full_name + "   private=" + $r.private + "   default_branch=" + $r.default_branch)
} catch {
  Write-Output "  FAILED to read repo"
  Show-Error $_.Exception
  exit 1
}

# ---- 3. invite ----
Write-Output ""
Write-Output "=== 3. invite collaborators (permission=$Permission) ==="
$ok = 0; $fail = 0
foreach ($u in $users) {
  $uri = "https://api.github.com/repos/$Owner/$Repo/collaborators/$u"
  $body = @{ permission = $Permission } | ConvertTo-Json -Compress
  try {
    $res = Invoke-RestMethod -Method Put -Uri $uri -Headers $headers -Proxy $proxy `
      -ContentType 'application/json' -Body $body -TimeoutSec 30
    $id = if ($res.id) { $res.id } else { ($res.invitee.id) }
    $login = if ($res.login) { $res.login } else { ($res.invitee.login) }
    Write-Output ("  OK      " + $u + "   -> id=" + $id + " login=" + $login)
    $ok++
  } catch {
    Write-Output ("  FAILED  " + $u)
    Show-Error $_.Exception
    $fail++
  }
  Start-Sleep -Milliseconds 600
}

# ---- 4. verify ----
Write-Output ""
Write-Output "=== 4. verify ==="
try {
  $c = Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/collaborators" -Headers $headers -Proxy $proxy -TimeoutSec 30
  Write-Output ("  collaborators now: " + $c.Count)
  foreach ($x in $c) {
    $perm = ''
    if ($x.permissions) {
      if ($x.permissions.admin) { $perm = 'admin' }
      elseif ($x.permissions.push) { $perm = 'write(push)' }
      elseif ($x.permissions.pull) { $perm = 'read(pull)' }
    }
    Write-Output ("    - " + $x.login + "   " + $perm)
  }
} catch {
  Write-Output "  could not list collaborators"
  Show-Error $_.Exception
}

Write-Output ""
Write-Output ("=== done: ok=" + $ok + " failed=" + $fail + " ===")
if ($fail -gt 0) { exit 1 }
