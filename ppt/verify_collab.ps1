# Verify collaborator invitations and permission level.
# Usage: .\verify_collab.ps1 -Token <PAT>
param([Parameter(Mandatory = $true)][string]$Token)

$ErrorActionPreference = 'Continue'
$proxy = 'http://127.0.0.1:14138'
$Owner = 'rrafiya'
$Repo = 'git_course'
$headers = @{
  'User-Agent'    = 'dsh'
  'Authorization' = "Bearer $Token"
  'Accept'        = 'application/vnd.github+json'
}

Write-Output "=== pending invitations ==="
try {
  $inv = Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/invitations" -Headers $headers -Proxy $proxy -TimeoutSec 30
  Write-Output ("  count: " + $inv.Count)
  foreach ($i in $inv) {
    $who = if ($i.invitee) { $i.invitee.login } else { $i.email }
    Write-Output ("    - " + $who + "   permission=" + $i.permissions + "   invited=" + $i.created_at)
  }
} catch { Write-Output ("  FAILED: " + $_.Exception.Message) }

Write-Output ""
Write-Output "=== direct collaborators (accepted) ==="
try {
  $c = Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/collaborators" -Headers $headers -Proxy $proxy -TimeoutSec 30
  Write-Output ("  count: " + $c.Count)
  foreach ($x in $c) {
    $perm = ''
    if ($x.permissions) {
      if ($x.permissions.admin) { $perm = 'admin' }
      elseif ($x.permissions.maintain) { $perm = 'maintain' }
      elseif ($x.permissions.push) { $perm = 'write' }
      elseif ($x.permissions.triage) { $perm = 'triage' }
      elseif ($x.permissions.pull) { $perm = 'read' }
    }
    Write-Output ("    - " + $x.login + "   role=" + $x.role_name + "   " + $perm)
  }
} catch { Write-Output ("  FAILED: " + $_.Exception.Message) }

Write-Output ""
Write-Output "=== permission lookup for each invitee ==="
foreach ($u in @('sunzhi-01', 'liulangaaa', 'lhm-11')) {
  try {
    $p = Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/collaborators/$u/permission" -Headers $headers -Proxy $proxy -TimeoutSec 25
    Write-Output ("  " + $u.PadRight(13) + " permission=" + $p.permission + "   role_name=" + $p.role_name + "   user=" + $p.user.login)
  } catch {
    $code = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { '?' }
    Write-Output ("  " + $u.PadRight(13) + " lookup HTTP " + $code + " (pending invitees may 404 here)")
  }
  Start-Sleep -Milliseconds 400
}
