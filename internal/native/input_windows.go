//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"syscall"
	"unsafe"
)

type mouseInput struct {
	DX, DY            int32
	Data, Flags, Time uint32
	Extra             uintptr
}
type inputEvent struct {
	Type  uint32
	Mouse mouseInput
}

func physicalClick(pid uint32, hwnd uintptr, control dispatch, child int) (any, error) {
	if err := verifyDesktop("Default"); err != nil {
		return nil, Fail("visible_host_required", "Physical mouse input requires native.start visible=true", nil)
	}
	awareness := user32.NewProc("SetThreadDpiAwarenessContext")
	previous, _, _ := awareness.Call(^uintptr(3))
	if previous != 0 {
		defer awareness.Call(previous)
	}
	var left, top, width, height int32
	value, err := makeVariant(child, nil)
	if err != nil {
		return nil, err
	}
	defer value.clear()
	method := (*[28]uintptr)(unsafe.Pointer(*(*uintptr)(unsafe.Pointer(control.ptr))))[22]
	hr, _, _ := syscall.SyscallN(method, control.ptr, uintptr(unsafe.Pointer(&left)), uintptr(unsafe.Pointer(&top)), uintptr(unsafe.Pointer(&width)), uintptr(unsafe.Pointer(&height)), uintptr(unsafe.Pointer(&value)))
	if failed(hr) || width <= 0 || height <= 0 {
		return nil, fmt.Errorf("control has no usable on-screen rectangle")
	}
	root, _, _ := user32.NewProc("GetAncestor").Call(hwnd, 2)
	user32.NewProc("SetForegroundWindow").Call(root)
	user32.NewProc("SetWindowPos").Call(root, 0, 0, 0, 0, 0, 0x43)
	x, y := left+width/2, top+height/2
	hit, _, _ := user32.NewProc("WindowFromPoint").Call(uintptr(uint64(uint32(x)) | uint64(uint32(y))<<32))
	if pidOf(hit) != pid {
		return nil, Fail("click_obstructed", "Another process covers the target control", map[string]any{"x": x, "y": y, "expected_pid": pid, "actual_pid": pidOf(hit), "actual_window": uint64(hit), "actual_title": textOf(hit)})
	}
	if ok, _, e := user32.NewProc("SetCursorPos").Call(uintptr(x), uintptr(y)); ok == 0 {
		return nil, winError("SetCursorPos", e)
	}
	inputs := []inputEvent{{Mouse: mouseInput{Flags: 2}}, {Mouse: mouseInput{Flags: 4}}}
	count, _, e := user32.NewProc("SendInput").Call(uintptr(len(inputs)), uintptr(unsafe.Pointer(&inputs[0])), unsafe.Sizeof(inputs[0]))
	if count != uintptr(len(inputs)) {
		return nil, winError("SendInput", e)
	}
	return map[string]any{"input_sent": true, "backend": "native SendInput", "hwnd": uint64(hwnd), "x": x, "y": y, "completion_observed": false}, nil
}
