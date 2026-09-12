//go:build darwin && (amd64 || arm64)

package darwinapi

func AECreateDesc() uintptr

//go:cgo_import_dynamic ww_AECreateDesc AECreateDesc "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AEDisposeDesc() uintptr

//go:cgo_import_dynamic ww_AEDisposeDesc AEDisposeDesc "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AECreateAppleEvent() uintptr

//go:cgo_import_dynamic ww_AECreateAppleEvent AECreateAppleEvent "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AEPutParamDesc() uintptr

//go:cgo_import_dynamic ww_AEPutParamDesc AEPutParamDesc "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AESendMessage() uintptr

//go:cgo_import_dynamic ww_AESendMessage AESendMessage "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AEGetParamDesc() uintptr

//go:cgo_import_dynamic ww_AEGetParamDesc AEGetParamDesc "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AEGetDescDataSize() uintptr

//go:cgo_import_dynamic ww_AEGetDescDataSize AEGetDescDataSize "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AEGetDescData() uintptr

//go:cgo_import_dynamic ww_AEGetDescData AEGetDescData "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AECreateList() uintptr

//go:cgo_import_dynamic ww_AECreateList AECreateList "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AEPutDesc() uintptr

//go:cgo_import_dynamic ww_AEPutDesc AEPutDesc "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AECountItems() uintptr

//go:cgo_import_dynamic ww_AECountItems AECountItems "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AEGetNthDesc() uintptr

//go:cgo_import_dynamic ww_AEGetNthDesc AEGetNthDesc "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AEDuplicateDesc() uintptr

//go:cgo_import_dynamic ww_AEDuplicateDesc AEDuplicateDesc "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func CreateObjSpecifier() uintptr

//go:cgo_import_dynamic ww_CreateObjSpecifier CreateObjSpecifier "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func AECoerceDesc() uintptr

//go:cgo_import_dynamic ww_AECoerceDesc AECoerceDesc "/System/Library/Frameworks/CoreServices.framework/Versions/A/CoreServices"

func ProcListAllPIDs() uintptr

//go:cgo_import_dynamic ww_ProcListAllPIDs proc_listallpids "/usr/lib/libproc.dylib"

func ProcPIDPath() uintptr

//go:cgo_import_dynamic ww_ProcPIDPath proc_pidpath "/usr/lib/libproc.dylib"
