package signing

import (
	"crypto/x509"
	"fmt"
	"strings"
	"syscall"
	"time"
	"unsafe"
)

// Query public certificate metadata without PowerShell startup or key access.
func existingCertificate(marker string) (*Certificate, error) {
	api := syscall.NewLazyDLL("crypt32.dll")
	name, _ := syscall.UTF16PtrFromString("My")
	store, _, err := api.NewProc("CertOpenStore").Call(10, 0, 0, 0x1c000, uintptr(unsafe.Pointer(name)))
	if store == 0 {
		return nil, err
	}
	defer api.NewProc("CertCloseStore").Call(store, 0)
	type certContext struct {
		Encoding    uint32
		Encoded     *byte
		Length      uint32
		Info, Store uintptr
	}
	property := api.NewProc("CertGetCertificateContextProperty")
	var previous uintptr
	var best *Certificate
	var bestExpiry time.Time
	for {
		certificate, _, _ := api.NewProc("CertEnumCertificatesInStore").Call(store, previous)
		previous = certificate
		if certificate == 0 {
			break
		}
		var size uint32
		ok, _, _ := property.Call(certificate, 11, 0, uintptr(unsafe.Pointer(&size)))
		if ok == 0 || size < 2 || size > 65536 {
			continue
		}
		friendly := make([]uint16, (size+1)/2)
		ok, _, _ = property.Call(certificate, 11, uintptr(unsafe.Pointer(&friendly[0])), uintptr(unsafe.Pointer(&size)))
		if ok == 0 || syscall.UTF16ToString(friendly) != marker {
			continue
		}
		ok, _, _ = property.Call(certificate, 2, 0, uintptr(unsafe.Pointer(&size)))
		if ok == 0 {
			continue
		}
		c := (*certContext)(unsafe.Pointer(certificate))
		if c.Length == 0 || c.Length > 1<<20 {
			continue
		}
		parsed, e := x509.ParseCertificate(unsafe.Slice(c.Encoded, int(c.Length)))
		if e != nil {
			continue
		}
		now := time.Now()
		if now.Before(parsed.NotBefore) || !parsed.NotAfter.After(now.Add(24*time.Hour)) || !parsed.NotAfter.After(bestExpiry) {
			continue
		}
		codeSigning := false
		for _, usage := range parsed.ExtKeyUsage {
			if usage == x509.ExtKeyUsageCodeSigning || usage == x509.ExtKeyUsageAny {
				codeSigning = true
			}
		}
		if !codeSigning {
			continue
		}
		hash := make([]byte, 20)
		size = 20
		ok, _, _ = property.Call(certificate, 3, uintptr(unsafe.Pointer(&hash[0])), uintptr(unsafe.Pointer(&size)))
		if ok == 0 {
			continue
		}
		best = &Certificate{Thumbprint: strings.ToUpper(fmt.Sprintf("%x", hash)), Subject: parsed.Subject.String(), Expires: parsed.NotAfter.UTC().Format(time.RFC3339Nano)}
		bestExpiry = parsed.NotAfter
	}
	return best, nil
}
