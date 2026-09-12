package inspect

import (
	"bytes"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	"image/png"
	"math"
	"path/filepath"
	"wordwright.local/internal/office"
	"wordwright.local/internal/project"
)

type ImageComparison struct {
	Width                    int     `json:"width"`
	Height                   int     `json:"height"`
	Pixels                   int64   `json:"pixels"`
	ChangedPixels            int64   `json:"changed_pixels"`
	ChangedFraction          float64 `json:"changed_fraction"`
	MeanAbsoluteChannelError float64 `json:"mean_absolute_channel_error"`
	MaximumChannelError      int     `json:"maximum_channel_error"`
	Tolerance                int     `json:"channel_tolerance"`
	Identical                bool    `json:"pixel_identical"`
	ReferenceSHA256          string  `json:"reference_sha256"`
	ActualSHA256             string  `json:"actual_sha256"`
	Difference               string  `json:"difference_image,omitempty"`
}

func Image(file string) (image.Image, []byte, string, error) {
	b, e := project.Read(filepath.Dir(file), filepath.Base(file))
	if e != nil {
		return nil, nil, "", e
	}
	config, typ, e := image.DecodeConfig(bytes.NewReader(b))
	if e != nil {
		return nil, nil, "", e
	}
	if config.Width <= 0 || config.Height <= 0 || int64(config.Width)*int64(config.Height) > 64000000 {
		return nil, nil, "", fmt.Errorf("image pixel budget exceeded")
	}
	im, _, e := image.Decode(bytes.NewReader(b))
	return im, b, typ, e
}
func Compare(reference, actual, output string, tolerance int) (*ImageComparison, error) {
	if tolerance < 0 || tolerance > 255 {
		return nil, fmt.Errorf("channel tolerance must be 0..255")
	}
	a, ab, _, e := Image(reference)
	if e != nil {
		return nil, e
	}
	b, bb, _, e := Image(actual)
	if e != nil {
		return nil, e
	}
	ra, rb := a.Bounds(), b.Bounds()
	if ra.Size() != rb.Size() {
		return nil, fmt.Errorf("image sizes differ (%v vs %v); no hidden rescaling or alignment was applied", ra.Size(), rb.Size())
	}
	r := &ImageComparison{Width: ra.Dx(), Height: ra.Dy(), Pixels: int64(ra.Dx()) * int64(ra.Dy()), Tolerance: tolerance, ReferenceSHA256: office.Hash(ab), ActualSHA256: office.Hash(bb)}
	delta := image.NewNRGBA(image.Rect(0, 0, r.Width, r.Height))
	sum := float64(0)
	for y := 0; y < r.Height; y++ {
		for x := 0; x < r.Width; x++ {
			ar, ag, abl, aa := a.At(ra.Min.X+x, ra.Min.Y+y).RGBA()
			br, bg, bbl, ba := b.At(rb.Min.X+x, rb.Min.Y+y).RGBA()
			ac := []uint32{ar, ag, abl, aa}
			bc := []uint32{br, bg, bbl, ba}
			changed := false
			for ch := 0; ch < 4; ch++ {
				d := int(math.Abs(float64(ac[ch])-float64(bc[ch]))/257 + 0.5)
				if d > r.MaximumChannelError {
					r.MaximumChannelError = d
				}
				sum += float64(d)
				if d > tolerance {
					changed = true
				}
				if ch < 3 {
					delta.Pix[delta.PixOffset(x, y)+ch] = byte(d)
				}
			}
			delta.Pix[delta.PixOffset(x, y)+3] = 255
			if changed {
				r.ChangedPixels++
			}
		}
	}
	r.Identical = r.MaximumChannelError == 0
	r.ChangedFraction = float64(r.ChangedPixels) / float64(r.Pixels)
	r.MeanAbsoluteChannelError = sum / float64(r.Pixels*4)
	if output != "" {
		var buf bytes.Buffer
		if e = png.Encode(&buf, delta); e != nil {
			return nil, e
		}
		if e = project.AtomicWrite(output, buf.Bytes()); e != nil {
			return nil, e
		}
		r.Difference = output
	}
	return r, nil
}
