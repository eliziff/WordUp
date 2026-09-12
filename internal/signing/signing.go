// Package signing uses Microsoft's SignTool and Office SIP, never exports keys,
// and publishes a new output only after its strongest (V3) VBA digest verifies.
package signing

import (
	"bytes"
	"context"
	"encoding/hex"
	"fmt"
	"net/url"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"time"

	"github.com/eliziff/WordUp/internal/office"
)

type Options struct {
	Thumbprint    string `json:"thumbprint"`
	SignTool      string `json:"signtool,omitempty"`
	TimestampURL  string `json:"timestamp_url,omitempty"`
	StoreLocation string `json:"store_location,omitempty"`
}

type Certificate struct {
	Thumbprint string `json:"thumbprint"`
	Subject    string `json:"subject"`
	Expires    string `json:"expires"`
	Created    bool   `json:"created"`
}

type Report struct {
	Artifact       string       `json:"artifact"`
	SHA256         string       `json:"sha256"`
	SourceSHA256   string       `json:"source_sha256,omitempty"`
	Thumbprint     string       `json:"thumbprint,omitempty"`
	Verified       bool         `json:"verified"`
	DigestVerified bool         `json:"digest_verified"`
	Signed         bool         `json:"signed"`
	TrustError     string       `json:"trust_error,omitempty"`
	Signatures     []string     `json:"signatures,omitempty"`
	Diagnostics    []string     `json:"diagnostics"`
	Certificate    *Certificate `json:"certificate,omitempty"`
}

func tool(configured string) (string, error) {
	if runtime.GOOS != "windows" {
		return "", fmt.Errorf("VBA signing requires Windows SignTool and the Microsoft Office SIP")
	}
	if configured != "" {
		return filepath.Abs(configured)
	}
	arch := runtime.GOARCH
	if arch == "amd64" {
		arch = "x64"
	}
	if arch == "386" {
		arch = "x86"
	}
	paths, _ := filepath.Glob(filepath.Join(os.Getenv("ProgramFiles(x86)"), "Windows Kits", "10", "bin", "*", arch, "signtool.exe"))
	sort.Sort(sort.Reverse(sort.StringSlice(paths)))
	if len(paths) > 0 {
		return paths[0], nil
	}
	return exec.LookPath("signtool.exe")
}

func command(ctx context.Context, executable string, args ...string) (string, error) {
	ctx, cancel := context.WithTimeout(ctx, 60*time.Second)
	defer cancel()
	c := exec.CommandContext(ctx, executable, args...)
	hide(c)
	b, err := c.CombinedOutput()
	if err != nil {
		return string(b), fmt.Errorf("Microsoft SignTool failed (Office SIP must be registered): %w: %s", err, b)
	}
	return string(b), nil
}

func Verify(ctx context.Context, file string, opts Options) (*Report, error) {
	executable, err := tool(opts.SignTool)
	if err != nil {
		return nil, err
	}
	b, err := os.ReadFile(file)
	if err != nil {
		return nil, err
	}
	p, err := office.ReadPackage(b)
	if err != nil {
		return nil, err
	}
	r := &Report{Artifact: file, SHA256: office.Hash(b)}
	for name := range p.Files {
		if strings.Contains(strings.ToLower(name), "vbaprojectsignature") && strings.HasSuffix(name, ".bin") {
			r.Signatures = append(r.Signatures, name)
		}
	}
	sort.Strings(r.Signatures)
	if len(r.Signatures) == 0 {
		return r, fmt.Errorf("no VBA project signature present")
	}
	worker, err := os.Executable()
	if err != nil {
		return r, err
	}
	if _, err = command(ctx, worker, "__verify_vba_digest", file); err != nil {
		return r, err
	}
	r.DigestVerified = true
	r.Signed = true
	diagnostic, err := command(ctx, executable, "verify", "/pa", "/v", file)
	r.Diagnostics = append(r.Diagnostics, diagnostic)
	// Office SIP validates the newest signature present; /all is NOT supported.
	r.Verified = err == nil
	if err != nil {
		r.TrustError = err.Error()
	}
	after, e := os.ReadFile(file)
	if e != nil || office.Hash(after) != r.SHA256 {
		r.DigestVerified = false
		r.Verified = false
		return r, fmt.Errorf("artifact changed during signature verification")
	}
	return r, nil
}

func Sign(ctx context.Context, input, output string, opts Options) (*Report, error) {
	ctx, cancel := context.WithTimeout(ctx, 60*time.Second)
	defer cancel()
	if opts.StoreLocation != "" && opts.StoreLocation != "current_user" && opts.StoreLocation != "local_machine" {
		return nil, fmt.Errorf("store_location must be current_user or local_machine")
	}
	thumb := strings.ReplaceAll(opts.Thumbprint, " ", "")
	decoded, err := hex.DecodeString(thumb)
	if opts.Thumbprint != "" && (err != nil || len(decoded) != 20) {
		return nil, fmt.Errorf("an exact certificate SHA-1 thumbprint (40 hexadecimal digits) is required")
	}
	if opts.TimestampURL != "" {
		u, e := url.Parse(opts.TimestampURL)
		if e != nil || u.Host == "" || (u.Scheme != "http" && u.Scheme != "https") {
			return nil, fmt.Errorf("timestamp_url must be HTTP(S)")
		}
	}
	executable, err := tool(opts.SignTool)
	if err != nil {
		return nil, err
	}
	if !strings.EqualFold(filepath.Ext(output), ".dotm") {
		return nil, fmt.Errorf("output must be a new .dotm file")
	}
	if _, err = os.Lstat(output); !os.IsNotExist(err) {
		return nil, fmt.Errorf("signing output must not already exist")
	}
	b, err := os.ReadFile(input)
	if err != nil {
		return nil, err
	}
	p, err := office.ReadPackage(b)
	if err != nil {
		return nil, err
	}
	if len(p.Files["word/vbaProject.bin"]) == 0 {
		return nil, fmt.Errorf("no VBA project")
	}
	if p.HasSignatures() {
		return nil, fmt.Errorf("build an unsigned candidate before signing; existing signatures are never silently removed")
	}
	var certificate *Certificate
	if thumb == "" {
		if opts.StoreLocation == "local_machine" {
			return nil, fmt.Errorf("automatic certificates use the current user store")
		}
		certificate, err = EnsureCertificate(ctx)
		if err != nil {
			return nil, err
		}
		thumb = certificate.Thumbprint
	}
	stage, err := os.MkdirTemp(filepath.Dir(output), ".wordup-sign-")
	if err != nil {
		return nil, err
	}
	defer os.RemoveAll(stage)
	file := filepath.Join(stage, "candidate.dotm")
	if err = os.WriteFile(file, b, 0600); err != nil {
		return nil, err
	}
	r := &Report{Artifact: output, SourceSHA256: office.Hash(b), Thumbprint: strings.ToUpper(thumb), Certificate: certificate}
	for _, version := range []string{"legacy", "agile", "v3"} {
		args := []string{"sign", "/sha1", thumb, "/s", "My", "/fd", "SHA256"}
		if opts.StoreLocation == "local_machine" {
			args = append(args, "/sm")
		}
		if opts.TimestampURL != "" {
			args = append(args, "/tr", opts.TimestampURL, "/td", "SHA256")
		}
		args = append(args, file)
		diagnostic, e := command(ctx, executable, args...)
		r.Diagnostics = append(r.Diagnostics, diagnostic)
		if e != nil {
			return r, e
		}
		r.Signatures = append(r.Signatures, version)
	}
	v, e := Verify(ctx, file, Options{SignTool: executable})
	if v != nil {
		r.Diagnostics = append(r.Diagnostics, v.Diagnostics...)
		r.DigestVerified = v.DigestVerified
		r.Verified = v.Verified
		r.TrustError = v.TrustError
	}
	if e != nil {
		return r, e
	}
	b, err = os.ReadFile(file)
	if err != nil {
		return r, err
	}
	signedPackage, err := office.ReadPackage(b)
	if err != nil {
		return r, err
	}
	if !bytes.Equal(p.Files["word/vbaProject.bin"], signedPackage.Files["word/vbaProject.bin"]) {
		return r, fmt.Errorf("signing changed the compiled VBA project; output not published")
	}
	// Exclusive creation prevents a concurrent writer from losing its artifact.
	out, err := os.OpenFile(output, os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0600)
	if err != nil {
		return r, err
	}
	_, err = out.Write(b)
	closeErr := out.Close()
	if err == nil {
		err = closeErr
	}
	if err != nil {
		_ = os.Remove(output)
		return r, err
	}
	r.SHA256 = office.Hash(b)
	r.Signed = true
	return r, nil
}
