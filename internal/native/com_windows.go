//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"math"
	"runtime"
	"sort"
	"syscall"
	"time"
	"unicode/utf16"
	"unsafe"
)

// Direct Automation ABI. No Python, COM interop runtime, injected DLL or VSTO.
var ole32 = syscall.NewLazyDLL("ole32.dll")
var oleaut = syscall.NewLazyDLL("oleaut32.dll")
var oleacc = syscall.NewLazyDLL("oleacc.dll")
var coInit = ole32.NewProc("CoInitializeEx")
var coUninit = ole32.NewProc("CoUninitialize")
var sysAlloc = oleaut.NewProc("SysAllocStringLen")
var sysFree = oleaut.NewProc("SysFreeString")
var sysLen = oleaut.NewProc("SysStringLen")
var variantClear = oleaut.NewProc("VariantClear")
var variantCopyInd = oleaut.NewProc("VariantCopyInd")
var decimalToString = oleaut.NewProc("VarBstrFromDec")
var stringToDecimal = oleaut.NewProc("VarDecFromStr")
var arrayCreate = oleaut.NewProc("SafeArrayCreateVector")
var arrayPut = oleaut.NewProc("SafeArrayPutElement")
var arrayGet = oleaut.NewProc("SafeArrayGetElement")
var arrayDestroy = oleaut.NewProc("SafeArrayDestroy")
var arrayDim = oleaut.NewProc("SafeArrayGetDim")
var arrayLower = oleaut.NewProc("SafeArrayGetLBound")
var arrayUpper = oleaut.NewProc("SafeArrayGetUBound")
var accessibleWindow = oleacc.NewProc("AccessibleObjectFromWindow")
var iidDispatch = syscall.GUID{Data1: 0x00020400, Data4: [8]byte{0xc0, 0, 0, 0, 0, 0, 0, 0x46}}
var iidNull = syscall.GUID{}

type dispatch struct{ ptr uintptr }
type variant struct {
	VT           uint16
	R1, R2, R3   uint16
	Value, Extra uint64
}
type dispParams struct {
	Args              *variant
	Named             *int32
	Count, NamedCount uint32
}
type exceptInfo struct {
	Code, Reserved                uint16
	Source, Description, HelpFile uintptr
	HelpContext                   uint32
	ReservedPtr, Deferred         uintptr
	Scode                         int32
}

func failed(hr uintptr) bool { return int32(hr) < 0 }
func (d dispatch) method(index int) uintptr {
	return (*[7]uintptr)(unsafe.Pointer(*(*uintptr)(unsafe.Pointer(d.ptr))))[index]
}
func (d dispatch) addRef() {
	if d.ptr != 0 {
		syscall.SyscallN(d.method(1), d.ptr)
	}
}
func (d dispatch) release() {
	if d.ptr != 0 {
		syscall.SyscallN(d.method(2), d.ptr)
	}
}
func (v *variant) clear() { variantClear.Call(uintptr(unsafe.Pointer(v))); *v = variant{} }
func bstr(s string) (uintptr, error) {
	a := utf16.Encode([]rune(s))
	var p uintptr
	if len(a) > 0 {
		p = uintptr(unsafe.Pointer(&a[0]))
	}
	b, _, _ := sysAlloc.Call(p, uintptr(len(a)))
	runtime.KeepAlive(a)
	if b == 0 && len(a) > 0 {
		return 0, fmt.Errorf("BSTR allocation failed")
	}
	return b, nil
}
func bstrText(p uintptr) (string, error) {
	if p == 0 {
		return "", nil
	}
	n, _, _ := sysLen.Call(p)
	if n > 16<<20 {
		return "", fmt.Errorf("BSTR exceeds 16-million-character response budget; request a smaller native range")
	}
	a := unsafe.Slice((*uint16)(unsafe.Pointer(p)), int(n))
	return string(utf16.Decode(a)), nil
}
func fromNativeWindow(hwnd uintptr, objid uint32) (dispatch, error) {
	var p uintptr
	hr, _, _ := accessibleWindow.Call(hwnd, uintptr(objid), uintptr(unsafe.Pointer(&iidDispatch)), uintptr(unsafe.Pointer(&p)))
	if failed(hr) || p == 0 {
		return dispatch{}, Fail("native_object_unavailable", fmt.Sprintf("AccessibleObjectFromWindow HRESULT 0x%08X", uint32(hr)), nil)
	}
	return dispatch{p}, nil
}
func (d dispatch) ids(names []string) ([]int32, error) {
	if d.ptr == 0 {
		return nil, fmt.Errorf("released COM object")
	}
	pointers := make([]*uint16, len(names))
	for i, s := range names {
		p, e := syscall.UTF16PtrFromString(s)
		if e != nil {
			return nil, e
		}
		pointers[i] = p
	}
	ids := make([]int32, len(names))
	var hr uintptr
	for attempt := 0; ; attempt++ {
		hr, _, _ = syscall.SyscallN(d.method(5), d.ptr, uintptr(unsafe.Pointer(&iidNull)), uintptr(unsafe.Pointer(&pointers[0])), uintptr(len(names)), 0, uintptr(unsafe.Pointer(&ids[0])))
		if !rejectedCall(hr) || attempt >= 7 {
			break
		}
		pump()
		time.Sleep(time.Duration(16<<min(attempt, 4)) * time.Millisecond)
	}
	runtime.KeepAlive(pointers)
	if failed(hr) {
		return nil, Fail("unknown_member", fmt.Sprintf("%s: HRESULT 0x%08X", names[0], uint32(hr)), nil)
	}
	return ids, nil
}
func makeVariant(value any, objects map[string]dispatch) (variant, error) {
	switch v := value.(type) {
	case nil:
		return variant{VT: 0}, nil
	case bool:
		if v {
			return variant{VT: 11, Value: 0xffff}, nil
		}
		return variant{VT: 11}, nil
	case string:
		b, e := bstr(v)
		return variant{VT: 8, Value: uint64(b)}, e
	case int:
		return makeVariant(float64(v), objects)
	case int64:
		return variant{VT: 20, Value: uint64(v)}, nil
	case float64:
		if math.IsNaN(v) || math.IsInf(v, 0) {
			return variant{}, fmt.Errorf("nonfinite Automation number")
		}
		if v == math.Trunc(v) && v >= math.MinInt32 && v <= math.MaxInt32 {
			return variant{VT: 3, Value: uint64(uint32(int32(v)))}, nil
		}
		return variant{VT: 5, Value: math.Float64bits(v)}, nil
	case dispatch:
		v.addRef()
		return variant{VT: 9, Value: uint64(v.ptr)}, nil
	case []any:
		if len(v) > 1000000 {
			return variant{}, fmt.Errorf("Automation array limit")
		}
		p, _, _ := arrayCreate.Call(12, 0, uintptr(len(v)))
		if p == 0 {
			return variant{}, fmt.Errorf("SAFEARRAY allocation failed")
		}
		for i, x := range v {
			elem, e := makeVariant(x, objects)
			if e != nil {
				arrayDestroy.Call(p)
				return variant{}, e
			}
			index := int32(i)
			hr, _, _ := arrayPut.Call(p, uintptr(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(&elem)))
			elem.clear()
			if failed(hr) {
				arrayDestroy.Call(p)
				return variant{}, fmt.Errorf("SAFEARRAY write failed")
			}
		}
		return variant{VT: 0x200c, Value: uint64(p)}, nil
	case map[string]any:
		if id, ok := v["object"].(string); ok {
			d, ok := objects[id]
			if !ok {
				return variant{}, fmt.Errorf("unknown object handle %s", id)
			}
			d.addRef()
			return variant{VT: 9, Value: uint64(d.ptr)}, nil
		}
		if m, ok := v["missing"].(bool); ok && m {
			return variant{VT: 10, Value: 0x80020004}, nil
		}
		if text, ok := v["int64"].(string); ok {
			var n int64
			if _, e := fmt.Sscan(text, &n); e != nil {
				return variant{}, e
			}
			return variant{VT: 20, Value: uint64(n)}, nil
		}
		return variant{}, fmt.Errorf("Automation maps must be object handles, missing arguments, or int64 values")
	default:
		return variant{}, fmt.Errorf("unsupported Automation input %T", value)
	}
}
func (d dispatch) invoke(member string, flags uint16, pos []any, named map[string]any, objects map[string]dispatch) (variant, error) {
	if len(pos)+len(named) > 128 {
		return variant{}, fmt.Errorf("Automation argument limit")
	}
	names := []string{member}
	keys := make([]string, 0, len(named))
	for k := range named {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	names = append(names, keys...)
	ids, e := d.ids(names)
	if e != nil {
		return variant{}, e
	}
	return d.invokeIDs(member, flags, pos, named, objects, keys, ids)
}

// IAccessible has standardized DISPIDs: callers can bypass repeated remote
// GetIDsOfNames without caching identities from mutable Word objects.
func (d dispatch) invokeIDs(member string, flags uint16, pos []any, named map[string]any, objects map[string]dispatch, keys []string, ids []int32) (variant, error) {
	args := make([]variant, 0, len(pos)+len(named))
	namedIDs := []int32{}
	cleanup := func() {
		for i := range args {
			args[i].clear()
		}
	}
	defer cleanup()
	for i, k := range keys {
		v, e := makeVariant(named[k], objects)
		if e != nil {
			return variant{}, e
		}
		args = append(args, v)
		namedIDs = append(namedIDs, ids[i+1])
	}
	for i := len(pos) - 1; i >= 0; i-- {
		v, e := makeVariant(pos[i], objects)
		if e != nil {
			return variant{}, e
		}
		args = append(args, v)
	}
	if flags&12 != 0 {
		if len(named) > 0 {
			return variant{}, fmt.Errorf("property puts use positional indexes followed by the value")
		}
		namedIDs = []int32{-3}
	}
	params := dispParams{Count: uint32(len(args)), NamedCount: uint32(len(namedIDs))}
	if len(args) > 0 {
		params.Args = &args[0]
	}
	if len(namedIDs) > 0 {
		params.Named = &namedIDs[0]
	}
	var result variant
	var ex exceptInfo
	var argError uint32
	var hr uintptr
	for attempt := 0; ; attempt++ {
		hr, _, _ = syscall.SyscallN(d.method(6), d.ptr, uintptr(uint32(ids[0])), uintptr(unsafe.Pointer(&iidNull)), 0, uintptr(flags), uintptr(unsafe.Pointer(&params)), uintptr(unsafe.Pointer(&result)), uintptr(unsafe.Pointer(&ex)), uintptr(unsafe.Pointer(&argError)))
		// Only rejection-before-execution is retried. Runtime errors and timeouts
		// are never retried: doing so could repeat a successful document mutation.
		if !rejectedCall(hr) || attempt >= 7 {
			break
		}
		result.clear()
		pump()
		time.Sleep(time.Duration(16<<min(attempt, 4)) * time.Millisecond)
	}
	runtime.KeepAlive(args)
	runtime.KeepAlive(namedIDs)
	if ex.Deferred != 0 {
		syscall.SyscallN(ex.Deferred, uintptr(unsafe.Pointer(&ex)))
	}
	source, sourceErr := bstrText(ex.Source)
	message, messageErr := bstrText(ex.Description)
	for _, p := range []uintptr{ex.Source, ex.Description, ex.HelpFile} {
		if p != 0 {
			sysFree.Call(p)
		}
	}
	if sourceErr != nil || messageErr != nil {
		result.clear()
		return variant{}, fmt.Errorf("COM exception text exceeded response budget (HRESULT 0x%08X)", uint32(hr))
	}
	if failed(hr) {
		result.clear()
		return variant{}, Fail("word_automation_error", fmt.Sprintf("%s: %s (HRESULT 0x%08X)", member, message, uint32(hr)), map[string]any{"member": member, "hresult": fmt.Sprintf("0x%08X", uint32(hr)), "source": source, "word_error": ex.Scode, "argument_index_reversed": argError})
	}
	return result, nil
}
func (d dispatch) get(name string, args ...any) (variant, error) {
	return d.invoke(name, 2, args, nil, nil)
}
func (d dispatch) call(name string, args ...any) (variant, error) {
	return d.invoke(name, 1, args, nil, nil)
}
func (d dispatch) put(name string, v any) error {
	r, e := d.invoke(name, 4, []any{v}, nil, nil)
	r.clear()
	return e
}
func (v *variant) object() (dispatch, error) {
	if v.VT != 9 || v.Value == 0 {
		return dispatch{}, fmt.Errorf("Automation result is not an object (VT=%d)", v.VT)
	}
	d := dispatch{uintptr(v.Value)}
	v.Value = 0
	v.VT = 0
	return d, nil
}
func (v *variant) value(depth int) (any, error) {
	if depth > 32 {
		return nil, fmt.Errorf("Automation nesting limit")
	}
	if v.VT&0x4000 != 0 {
		var copied variant
		hr, _, _ := variantCopyInd.Call(uintptr(unsafe.Pointer(&copied)), uintptr(unsafe.Pointer(v)))
		if failed(hr) {
			return nil, fmt.Errorf("VariantCopyInd: 0x%08X", uint32(hr))
		}
		defer copied.clear()
		return copied.value(depth + 1)
	}
	if v.VT&0x2000 != 0 {
		return v.arrayValue(depth)
	}

	switch v.VT {
	case 0, 1:
		return nil, nil
	case 2:
		return int16(v.Value), nil
	case 3:
		return int32(v.Value), nil
	case 4:
		return math.Float32frombits(uint32(v.Value)), nil
	case 5:
		return math.Float64frombits(v.Value), nil
	case 6:
		return map[string]any{"currency_scaled_10000": fmt.Sprint(int64(v.Value))}, nil
	case 7:
		return map[string]any{"automation_date": math.Float64frombits(v.Value)}, nil
	case 8:
		return bstrText(uintptr(v.Value))
	case 10:
		return map[string]any{"error": fmt.Sprintf("0x%08X", uint32(v.Value))}, nil
	case 11:
		return int16(v.Value) != 0, nil
	case 16:
		return int8(v.Value), nil
	case 17:
		return uint8(v.Value), nil
	case 18:
		return uint16(v.Value), nil
	case 19, 23:
		return uint32(v.Value), nil
	case 20:
		return map[string]any{"int64": fmt.Sprint(int64(v.Value))}, nil
	case 21:
		return map[string]any{"uint64": fmt.Sprint(v.Value)}, nil
	case 22:
		return int32(v.Value), nil
	case 14:
		var text uintptr
		hr, _, _ := decimalToString.Call(uintptr(unsafe.Pointer(v)), 0x409, 0, uintptr(unsafe.Pointer(&text)))
		if failed(hr) {
			return nil, fmt.Errorf("DECIMAL conversion failed")
		}
		defer sysFree.Call(text)
		value, err := bstrText(text)
		return map[string]any{"decimal": value}, err
	case 9, 13:
		return nil, fmt.Errorf("save this object result with an 'as' handle")
	default:
		return nil, fmt.Errorf("unimplemented Automation variant type %d; not coerced", v.VT)
	}
}

func rejectedCall(hr uintptr) bool { return uint32(hr) == 0x80010001 || uint32(hr) == 0x8001010a }
func (v *variant) arrayValue(depth int) (any, error) {
	p := uintptr(v.Value)
	dim, _, _ := arrayDim.Call(p)
	if dim < 1 || dim > 16 {
		return nil, fmt.Errorf("Automation array dimensionality budget")
	}
	type bounds struct {
		Lower  int32 `json:"lower"`
		Length int64 `json:"length"`
	}
	shape := make([]bounds, dim)
	index := make([]int32, dim)
	total := int64(1)
	for i := range shape {
		var hi, lo int32
		a, _, _ := arrayLower.Call(p, uintptr(i+1), uintptr(unsafe.Pointer(&lo)))
		b, _, _ := arrayUpper.Call(p, uintptr(i+1), uintptr(unsafe.Pointer(&hi)))
		n := int64(hi) - int64(lo) + 1
		if failed(a) || failed(b) || n < 0 || n > 1000000 || n != 0 && total > 1000000/n {
			return nil, fmt.Errorf("Automation array size budget")
		}
		shape[i] = bounds{lo, n}
		index[i] = lo
		total *= n
	}
	base := v.VT & 0xfff
	out := make([]any, 0, total)
	for n := int64(0); n < total; n++ {
		var elem variant
		var target unsafe.Pointer
		if base == 12 {
			target = unsafe.Pointer(&elem)
		} else {
			elem.VT = base
			target = unsafe.Pointer(&elem.Value)
			if base == 14 {
				target = unsafe.Pointer(&elem)
			}
		}
		hr, _, _ := arrayGet.Call(p, uintptr(unsafe.Pointer(&index[0])), uintptr(target))
		if failed(hr) {
			elem.clear()
			return nil, fmt.Errorf("SafeArrayGetElement: 0x%08X", uint32(hr))
		}
		if base == 14 {
			elem.VT = 14
		}
		value, e := elem.value(depth + 1)
		elem.clear()
		if e != nil {
			return nil, e
		}
		out = append(out, value)
		// COM SAFEARRAY index 0 varies fastest; preserve all original lower bounds.
		for j := 0; j < len(index); j++ {
			index[j]++
			if int64(index[j])-int64(shape[j].Lower) < shape[j].Length {
				break
			}
			index[j] = shape[j].Lower
		}
	}
	return map[string]any{"array": out, "bounds": shape, "element_vartype": base, "order": "first-dimension-fastest"}, nil
}
