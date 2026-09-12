//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"image"
	"image/png"
	"os"
	"path/filepath"
	"runtime"
	"syscall"
	"unsafe"

	"wordwright.local/internal/office"
)

var gdi = syscall.NewLazyDLL("gdi32.dll")
var createDC = gdi.NewProc("CreateCompatibleDC")
var deleteDC = gdi.NewProc("DeleteDC")
var createDIB = gdi.NewProc("CreateDIBSection")
var selectObject = gdi.NewProc("SelectObject")
var deleteObject = gdi.NewProc("DeleteObject")
var setEMF = gdi.NewProc("SetEnhMetaFileBits")
var deleteEMF = gdi.NewProc("DeleteEnhMetaFile")
var playEMF = gdi.NewProc("PlayEnhMetaFile")
var printWindow = user32.NewProc("PrintWindow")
var getWindowRect = user32.NewProc("GetWindowRect")
var arrayAccess = oleaut.NewProc("SafeArrayAccessData")
var arrayUnaccess = oleaut.NewProc("SafeArrayUnaccessData")

type rect struct{ Left, Top, Right, Bottom int32 }
type bitmapHeader struct {
	Size                   uint32
	Width, Height          int32
	Planes, BitCount       uint16
	Compression, SizeImage uint32
	XPels, YPels           int32
	ClrUsed, ClrImportant  uint32
}
type dibSurface struct {
	dc, bitmap, prior uintptr
	bits              unsafe.Pointer
	width, height     int
}

func newSurface(w, h int) (*dibSurface, error) {
	if w <= 0 || h <= 0 || w > 20000 || h > 20000 || int64(w)*int64(h) > 64000000 {
		return nil, fmt.Errorf("render dimensions exceed pixel budget")
	}
	dc, _, e := createDC.Call(0)
	if dc == 0 {
		return nil, winError("CreateCompatibleDC", e)
	}
	hdr := bitmapHeader{Size: 40, Width: int32(w), Height: -int32(h), Planes: 1, BitCount: 32}
	var bits unsafe.Pointer
	bm, _, e := createDIB.Call(dc, uintptr(unsafe.Pointer(&hdr)), 0, uintptr(unsafe.Pointer(&bits)), 0, 0)
	if bm == 0 {
		deleteDC.Call(dc)
		return nil, winError("CreateDIBSection", e)
	}
	prior, _, e := selectObject.Call(dc, bm)
	if prior == 0 || prior == ^uintptr(0) {
		deleteObject.Call(bm)
		deleteDC.Call(dc)
		return nil, winError("SelectObject", e)
	}
	s := &dibSurface{dc, bm, prior, bits, w, h}
	b := unsafe.Slice((*byte)(bits), w*h*4)
	for i := range b {
		b[i] = 255
	}
	return s, nil
}
func (s *dibSurface) close() {
	if s == nil {
		return
	}
	selectObject.Call(s.dc, s.prior)
	deleteObject.Call(s.bitmap)
	deleteDC.Call(s.dc)
}
func (s *dibSurface) write(file string) error {
	pixels := unsafe.Slice((*byte)(s.bits), s.width*s.height*4)
	im := image.NewNRGBA(image.Rect(0, 0, s.width, s.height))
	for i := 0; i < len(pixels); i += 4 {
		im.Pix[i] = pixels[i+2]
		im.Pix[i+1] = pixels[i+1]
		im.Pix[i+2] = pixels[i]
		im.Pix[i+3] = 255
	}
	if e := os.MkdirAll(filepath.Dir(file), 0700); e != nil {
		return e
	}
	f, e := os.Create(file)
	if e != nil {
		return e
	}
	e = (&png.Encoder{CompressionLevel: png.BestSpeed}).Encode(f, im)
	closeError := f.Close()
	if e != nil {
		return e
	}
	return closeError
}
func byteArray(v *variant) ([]byte, error) {
	if v.VT != 0x2011 {
		return nil, fmt.Errorf("expected SAFEARRAY(UI1), received VARIANT %d", v.VT)
	}
	p := uintptr(v.Value)
	dim, _, _ := arrayDim.Call(p)
	if dim != 1 {
		return nil, fmt.Errorf("expected one-dimensional picture bytes")
	}
	var lo, hi int32
	r1, _, _ := arrayLower.Call(p, 1, uintptr(unsafe.Pointer(&lo)))
	r2, _, _ := arrayUpper.Call(p, 1, uintptr(unsafe.Pointer(&hi)))
	n := int64(hi) - int64(lo) + 1
	if failed(r1) || failed(r2) || n < 0 || n > office.Limit {
		return nil, fmt.Errorf("invalid picture byte bounds")
	}
	if n == 0 {
		return []byte{}, nil
	}
	var data unsafe.Pointer
	hr, _, _ := arrayAccess.Call(p, uintptr(unsafe.Pointer(&data)))
	if failed(hr) || data == nil {
		return nil, fmt.Errorf("cannot access picture array")
	}
	defer arrayUnaccess.Call(p)
	return append([]byte(nil), unsafe.Slice((*byte)(data), int(n))...), nil
}
func renderEMF(data []byte, w, h int, file string) error {
	if len(data) == 0 || len(data) > office.Limit {
		return fmt.Errorf("empty or oversized EMF")
	}
	emf, _, e := setEMF.Call(uintptr(len(data)), uintptr(unsafe.Pointer(&data[0])))
	runtime.KeepAlive(data)
	if emf == 0 {
		return winError("SetEnhMetaFileBits", e)
	}
	defer deleteEMF.Call(emf)
	s, e := newSurface(w, h)
	if e != nil {
		return e
	}
	defer s.close()
	box := rect{Right: int32(w), Bottom: int32(h)}
	ok, _, e := playEMF.Call(s.dc, emf, uintptr(unsafe.Pointer(&box)))
	if ok == 0 {
		return winError("PlayEnhMetaFile", e)
	}
	return s.write(file)
}
func screenshot(hwnd uintptr, file string) error {
	var r rect
	ok, _, e := getWindowRect.Call(hwnd, uintptr(unsafe.Pointer(&r)))
	if ok == 0 {
		return winError("GetWindowRect", e)
	}
	s, e := newSurface(int(r.Right-r.Left), int(r.Bottom-r.Top))
	if e != nil {
		return e
	}
	defer s.close()
	ok, _, e = printWindow.Call(hwnd, s.dc, 2)
	if ok == 0 {
		return Fail("native_capture_unavailable", "The owned window did not paint through PrintWindow; no simulated screenshot was substituted", e.Error())
	}
	return s.write(file)
}
func objectProperty(d dispatch, name string, args ...any) (dispatch, error) {
	v, e := d.get(name, args...)
	if e != nil {
		return dispatch{}, e
	}
	defer v.clear()
	return v.object()
}
func scalarNumber(d dispatch, name string) (float64, error) {
	v, e := d.get(name)
	if e != nil {
		return 0, e
	}
	defer v.clear()
	x, e := v.value(0)
	if e != nil {
		return 0, e
	}
	switch n := x.(type) {
	case int32:
		return float64(n), nil
	case float64:
		return n, nil
	case float32:
		return float64(n), nil
	default:
		return 0, fmt.Errorf("%s is not numeric", name)
	}
}
func (h *wordHost) render(op Operation) (any, error) {
	doc, e := h.object(op.Target)
	if e != nil {
		return nil, e
	}
	if op.Target == "" || op.Target == "app" {
		return nil, fmt.Errorf("render requires an explicit document object handle")
	}
	if op.File == "" {
		return nil, fmt.Errorf("render output path required")
	}
	v, e := doc.call("Repaginate")
	v.clear()
	if e != nil {
		return nil, e
	}
	if op.Format == "pdf" {
		v, e = doc.invoke("ExportAsFixedFormat", 1, nil, map[string]any{"OutputFileName": op.File, "ExportFormat": 17, "OpenAfterExport": false}, h.objects)
		v.clear()
		if e != nil {
			return nil, e
		}
		return map[string]any{"path": op.File, "renderer": "Microsoft Word", "format": "pdf"}, nil
	}
	win, e := objectProperty(doc, "ActiveWindow")
	if e != nil {
		return nil, e
	}
	defer win.release()
	view, e := objectProperty(win, "View")
	if e != nil {
		return nil, e
	}
	if e = view.put("Type", 3); e != nil {
		view.release()
		return nil, e
	}
	view.release()
	panes, e := objectProperty(win, "Panes")
	if e != nil {
		return nil, e
	}
	defer panes.release()
	pane, e := objectProperty(panes, "Item", 1)
	if e != nil {
		return nil, e
	}
	defer pane.release()
	pages, e := objectProperty(pane, "Pages")
	if e != nil {
		return nil, e
	}
	defer pages.release()
	n, e := scalarNumber(pages, "Count")
	if e != nil {
		return nil, e
	}
	if n < 1 || n > 10000 {
		return nil, fmt.Errorf("page count outside render budget")
	}
	dpi := 144.0
	if x, ok := op.Value.(float64); ok {
		dpi = x
	}
	if dpi < 36 || dpi > 600 {
		return nil, fmt.Errorf("render DPI must be 36..600")
	}
	first, last := 1, int(n)
	if op.Child > 0 {
		first, last = op.Child, op.Child
		if first > int(n) {
			return nil, fmt.Errorf("page index outside document")
		}
	}
	out := []any{}
	for i := first; i <= last; i++ {
		page, e := objectProperty(pages, "Item", i)
		if e != nil {
			return nil, e
		}
		width, errW := scalarNumber(page, "Width")
		height, errH := scalarNumber(page, "Height")
		bits, e := page.get("EnhMetaFileBits")
		page.release()
		if errW != nil {
			return nil, errW
		}
		if errH != nil {
			return nil, errH
		}
		if e != nil {
			return nil, e
		}
		data, e := byteArray(&bits)
		bits.clear()
		if e != nil {
			return nil, e
		}
		path := filepath.Join(op.File, fmt.Sprintf("page-%04d.png", i))
		w, h := int(width*dpi/72+0.5), int(height*dpi/72+0.5)
		if e = renderEMF(data, w, h, path); e != nil {
			return nil, e
		}
		b, e := os.ReadFile(path)
		if e != nil {
			return nil, e
		}
		out = append(out, map[string]any{"page": i, "path": path, "width": w, "height": h, "sha256": office.Hash(b)})
	}
	return map[string]any{"renderer": "Microsoft Word Page.EnhMetaFileBits + Windows GDI", "dpi": dpi, "pages": out, "page_count": int(n)}, nil
}
