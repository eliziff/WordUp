package office

import (
	"bytes"
	"encoding/base64"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
)

// Store artwork in the control itself, with no runtime path or extraction step.
func (r *formRecord) applyPicture(value any) error {
	s, ok := value.(string)
	if !ok || len(s) > 24<<20 {
		return fmt.Errorf("PictureBase64 requires a base64 JPEG/GIF of at most 16 MiB")
	}
	raw, err := base64.StdEncoding.DecodeString(s)
	if err != nil {
		return fmt.Errorf("PictureBase64: %w", err)
	}
	if len(raw) > 16<<20 {
		return fmt.Errorf("picture exceeds 16 MiB")
	}
	if len(raw) > 0 {
		cfg, format, err := image.DecodeConfig(bytes.NewReader(raw))
		if err != nil || (format != "jpeg" && format != "gif") {
			return fmt.Errorf("PictureBase64 requires JPEG or GIF")
		}
		if cfg.Width <= 0 || cfg.Height <= 0 || int64(cfg.Width)*int64(cfg.Height) > 64000000 {
			return fmt.Errorf("picture exceeds pixel budget")
		}
	}
	p := 0
	for _, blob := range r.spec.blobs {
		end := p
		if r.mask&(1<<blob.bit) != 0 {
			end, err = pictureEnd(r.tail, p)
			if err != nil {
				return err
			}
		}
		if blob.name == "Picture" {
			var picture []byte
			if len(raw) > 0 {
				picture = append([]byte{0x04, 0x52, 0xe3, 0x0b, 0x91, 0x8f, 0xce, 0x11, 0x9d, 0xe3, 0, 0xaa, 0, 0x4b, 0xb8, 0x51}, dword(0x746c)...)
				picture = append(picture, dword(uint32(len(raw)))...)
				picture = append(picture, raw...)
				r.mask |= 1 << blob.bit
				r.values["Picture"] = 0xffff
			} else {
				r.mask &^= 1 << blob.bit
				delete(r.values, "Picture")
			}
			r.tail = append(append(append([]byte(nil), r.tail[:p]...), picture...), r.tail[end:]...)
			r.dirty = true
			return nil
		}
		p = end
	}
	return fmt.Errorf("%s has no supported control Picture stream", r.spec.name)
}
