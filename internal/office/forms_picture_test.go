package office

import (
	"bytes"
	"encoding/base64"
	"image"
	"image/jpeg"
	"testing"
)

func TestEmbeddedPicturePreservesFontOnReplaceAndRemove(t *testing.T) {
	var b bytes.Buffer
	if err := jpeg.Encode(&b, image.NewRGBA(image.Rect(0, 0, 2, 2)), nil); err != nil {
		t.Fatal(err)
	}
	for _, kind := range []string{"Image", "Label", "CommandButton"} {
		r, err := defaultControl(kind, "Art")
		if err != nil {
			t.Fatal(err)
		}
		font := append([]byte(nil), r.tail...)
		for i := 0; i < 2; i++ {
			if err := r.applyPicture(base64.StdEncoding.EncodeToString(b.Bytes())); err != nil {
				t.Fatal(err)
			}
			end, err := pictureEnd(r.tail, 0)
			if err != nil || !bytes.Equal(r.tail[24:end], b.Bytes()) || !bytes.Equal(r.tail[end:], font) {
				t.Fatalf("%s lost picture/font", kind)
			}
			if _, err := r.bytes(1252); err != nil {
				t.Fatal(err)
			}
		}
		if err := r.applyPicture(""); err != nil || !bytes.Equal(r.tail, font) {
			t.Fatal("removal changed font", err)
		}
		if err := r.applyPicture("bm90LWFuLWltYWdl"); err == nil {
			t.Fatal("accepted invalid image")
		}
	}
}
