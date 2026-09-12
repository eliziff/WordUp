//go:build windows

package signing

import (
	"context"
	"fmt"
	"github.com/eliziff/WordUp/internal/example"
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"os/exec"
	"path/filepath"
	"testing"
	"time"
)

func TestMain(m *testing.M) {
	if len(os.Args) == 3 && os.Args[1] == "__verify_vba_digest" {
		if e := VerifyDigestWorker(os.Args[2]); e != nil {
			fmt.Fprintln(os.Stderr, e)
			os.Exit(1)
		}
		return
	}
	os.Exit(m.Run())
}

func TestRejectInvalidThumbprint(t *testing.T) {
	if _, e := Sign(context.Background(), "missing.dotm", "out.dotm", Options{Thumbprint: "/a"}); e == nil {
		t.Fatal("invalid certificate selector accepted")
	}
}

func TestNativeSigning(t *testing.T) {
	if os.Getenv("WORDUP_SIGNING_TEST") != "1" {
		t.Skip("set WORDUP_SIGNING_TEST=1 to create and remove a disposable certificate")
	}
	marker := fmt.Sprintf("WordUp test %d", time.Now().UnixNano())
	certificate, e := ensureCertificate(context.Background(), marker)
	if e != nil {
		t.Fatal(e)
	}
	thumb := certificate.Thumbprint
	t.Cleanup(func() {
		c := exec.Command("powershell.exe", "-NoProfile", "-NonInteractive", "-Command", `$ErrorActionPreference='Stop'; Remove-Item -LiteralPath ('Cert:\CurrentUser\My\'+$env:WORDUP_TEST_CERT) -DeleteKey`)
		c.Env = append(os.Environ(), "WORDUP_TEST_CERT="+thumb)
		if b, e := c.CombinedOutput(); e != nil {
			t.Errorf("test certificate cleanup: %s %v", b, e)
		}
	})
	reused, e := ensureCertificate(context.Background(), marker)
	if e != nil || reused.Thumbprint != thumb || reused.Created || !certificate.Created {
		t.Fatalf("automatic creation/reuse: %v %v", reused, e)
	}
	root := t.TempDir()
	build, e := example.Studio(filepath.Join(root, "studio"))
	if e != nil {
		t.Fatal(e)
	}
	before, e := os.ReadFile(build.Artifact)
	if e != nil {
		t.Fatal(e)
	}
	output := filepath.Join(root, "signed.dotm")
	r, e := Sign(context.Background(), build.Artifact, output, Options{Thumbprint: thumb})
	if e != nil {
		t.Fatal(e)
	}
	if !r.Signed || !r.DigestVerified || len(r.Signatures) != 3 {
		t.Fatalf("incomplete signing: %+v", r)
	}
	after, _ := os.ReadFile(build.Artifact)
	if office.Hash(before) != office.Hash(after) {
		t.Fatal("source changed")
	}
	if _, e = Sign(context.Background(), build.Artifact, output, Options{Thumbprint: thumb}); e == nil {
		t.Fatal("existing output overwritten")
	}
	b, e := os.ReadFile(output)
	if e != nil {
		t.Fatal(e)
	}
	p, e := office.ReadPackage(b)
	if e != nil {
		t.Fatal(e)
	}
	if len(p.Files["word/vbaProjectSignatureV3.bin"]) == 0 {
		t.Fatal("V3 signature absent")
	}
	// Replace the VBA project with a distinct valid project, keeping signatures.
	v := office.NewVBA("ChangedProject")
	p.Files["word/vbaProject.bin"], e = v.Rewrite([]office.Module{{Name: "Main", Kind: "standard", Source: "Public Function Different() As Long\nDifferent=12345\nEnd Function\n"}}, nil)
	if e != nil {
		t.Fatal(e)
	}
	corrupt, e := p.Bytes()
	if e != nil {
		t.Fatal(e)
	}
	bad := filepath.Join(root, "tampered.dotm")
	if e = os.WriteFile(bad, corrupt, 0600); e != nil {
		t.Fatal(e)
	}
	if _, e = Verify(context.Background(), bad, Options{}); e == nil {
		t.Fatal("tampered VBA accepted")
	}
	t.Logf("signed legacy/agile/V3, digest verified, publisher trusted=%v, tampering rejected", r.Verified)
}
