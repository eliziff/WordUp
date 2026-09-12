package signing

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

// Windows creates and retains a non-exportable key in the current user's store.
// The marker identifies certificates created by this app, never a user's
// unrelated publisher identity. Expired or missing-key certificates are renewed.
func EnsureCertificate(ctx context.Context) (*Certificate, error) {
	return ensureCertificate(ctx, "WordUp local VBA signing")
}

func ensureCertificate(ctx context.Context, marker string) (*Certificate, error) {
	if existing, err := existingCertificate(marker); err != nil || existing != nil {
		return existing, err
	}
	input, _ := json.Marshal(map[string]string{"marker": marker})
	script := `$ErrorActionPreference='Stop'
$request=[Console]::In.ReadToEnd() | ConvertFrom-Json
$sid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$mutex=[Threading.Mutex]::new($false,('Local\WordUp.Certificate.'+$sid))
$locked=$false
try {
  try { $locked=$mutex.WaitOne(30000) } catch [Threading.AbandonedMutexException] { $locked=$true }
  if (!$locked) { throw 'Certificate creation is busy' }
  $now=Get-Date
  $certificate=Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object { $_.FriendlyName -eq $request.marker -and $_.HasPrivateKey -and $_.NotBefore -le $now -and $_.NotAfter -gt $now.AddDays(1) } | Sort-Object NotAfter -Descending | Select-Object -First 1
  $created=$false
  if (!$certificate) {
    $name=[Environment]::UserName -replace '[,+=<>#;"\\]','_'
    $certificate=New-SelfSignedCertificate -Type CodeSigningCert -Subject ('CN=WordUp '+$name) -FriendlyName $request.marker -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy NonExportable -NotAfter $now.AddYears(5)
    $created=$true
  }
  @{thumbprint=$certificate.Thumbprint;subject=$certificate.Subject;expires=$certificate.NotAfter.ToUniversalTime().ToString('o');created=$created} | ConvertTo-Json -Compress
} finally { if ($locked) { $mutex.ReleaseMutex() }; $mutex.Dispose() }`
	exe := filepath.Join(os.Getenv("SystemRoot"), "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
	c := exec.CommandContext(ctx, exe, "-NoLogo", "-NoProfile", "-NonInteractive", "-Command", script)
	hide(c)
	c.Stdin = strings.NewReader(string(input))
	b, err := c.CombinedOutput()
	if err != nil {
		return nil, fmt.Errorf("create/reuse local VBA certificate: %w: %s", err, b)
	}
	var certificate Certificate
	if err = json.Unmarshal(b, &certificate); err != nil {
		return nil, fmt.Errorf("certificate result: %w: %s", err, b)
	}
	return &certificate, nil
}
