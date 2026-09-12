//go:build darwin && arm64

#include "textflag.h"

TEXT ·AECreateDesc(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AECreateDesc(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AECreateDesc(SB),NOSPLIT,$0-0
 JMP ww_AECreateDesc(SB)

TEXT ·AEDisposeDesc(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AEDisposeDesc(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AEDisposeDesc(SB),NOSPLIT,$0-0
 JMP ww_AEDisposeDesc(SB)

TEXT ·AECreateAppleEvent(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AECreateAppleEvent(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AECreateAppleEvent(SB),NOSPLIT,$0-0
 JMP ww_AECreateAppleEvent(SB)

TEXT ·AEPutParamDesc(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AEPutParamDesc(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AEPutParamDesc(SB),NOSPLIT,$0-0
 JMP ww_AEPutParamDesc(SB)

TEXT ·AESendMessage(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AESendMessage(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AESendMessage(SB),NOSPLIT,$0-0
 JMP ww_AESendMessage(SB)

TEXT ·AEGetParamDesc(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AEGetParamDesc(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AEGetParamDesc(SB),NOSPLIT,$0-0
 JMP ww_AEGetParamDesc(SB)

TEXT ·AEGetDescDataSize(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AEGetDescDataSize(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AEGetDescDataSize(SB),NOSPLIT,$0-0
 JMP ww_AEGetDescDataSize(SB)

TEXT ·AEGetDescData(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AEGetDescData(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AEGetDescData(SB),NOSPLIT,$0-0
 JMP ww_AEGetDescData(SB)

TEXT ·AECreateList(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AECreateList(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AECreateList(SB),NOSPLIT,$0-0
 JMP ww_AECreateList(SB)

TEXT ·AEPutDesc(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AEPutDesc(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AEPutDesc(SB),NOSPLIT,$0-0
 JMP ww_AEPutDesc(SB)

TEXT ·AECountItems(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AECountItems(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AECountItems(SB),NOSPLIT,$0-0
 JMP ww_AECountItems(SB)

TEXT ·AEGetNthDesc(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AEGetNthDesc(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AEGetNthDesc(SB),NOSPLIT,$0-0
 JMP ww_AEGetNthDesc(SB)

TEXT ·AEDuplicateDesc(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AEDuplicateDesc(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AEDuplicateDesc(SB),NOSPLIT,$0-0
 JMP ww_AEDuplicateDesc(SB)

TEXT ·CreateObjSpecifier(SB),NOSPLIT,$0-8
 MOVD $·trampoline_CreateObjSpecifier(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_CreateObjSpecifier(SB),NOSPLIT,$0-0
 JMP ww_CreateObjSpecifier(SB)

TEXT ·AECoerceDesc(SB),NOSPLIT,$0-8
 MOVD $·trampoline_AECoerceDesc(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_AECoerceDesc(SB),NOSPLIT,$0-0
 JMP ww_AECoerceDesc(SB)


TEXT ·ProcListAllPIDs(SB),NOSPLIT,$0-8
 MOVD $·trampoline_ProcListAllPIDs(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_ProcListAllPIDs(SB),NOSPLIT,$0-0
 JMP ww_ProcListAllPIDs(SB)

TEXT ·ProcPIDPath(SB),NOSPLIT,$0-8
 MOVD $·trampoline_ProcPIDPath(SB), R0
 MOVD R0, ret+0(FP)
 RET
TEXT ·trampoline_ProcPIDPath(SB),NOSPLIT,$0-0
 JMP ww_ProcPIDPath(SB)
