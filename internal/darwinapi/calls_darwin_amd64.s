//go:build darwin && amd64

#include "textflag.h"

TEXT ·AECreateDesc(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AECreateDesc(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AECreateDesc(SB),NOSPLIT,$0-0
 JMP ww_AECreateDesc(SB)

TEXT ·AEDisposeDesc(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AEDisposeDesc(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AEDisposeDesc(SB),NOSPLIT,$0-0
 JMP ww_AEDisposeDesc(SB)

TEXT ·AECreateAppleEvent(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AECreateAppleEvent(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AECreateAppleEvent(SB),NOSPLIT,$0-0
 JMP ww_AECreateAppleEvent(SB)

TEXT ·AEPutParamDesc(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AEPutParamDesc(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AEPutParamDesc(SB),NOSPLIT,$0-0
 JMP ww_AEPutParamDesc(SB)

TEXT ·AESendMessage(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AESendMessage(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AESendMessage(SB),NOSPLIT,$0-0
 JMP ww_AESendMessage(SB)

TEXT ·AEGetParamDesc(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AEGetParamDesc(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AEGetParamDesc(SB),NOSPLIT,$0-0
 JMP ww_AEGetParamDesc(SB)

TEXT ·AEGetDescDataSize(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AEGetDescDataSize(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AEGetDescDataSize(SB),NOSPLIT,$0-0
 JMP ww_AEGetDescDataSize(SB)

TEXT ·AEGetDescData(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AEGetDescData(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AEGetDescData(SB),NOSPLIT,$0-0
 JMP ww_AEGetDescData(SB)

TEXT ·AECreateList(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AECreateList(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AECreateList(SB),NOSPLIT,$0-0
 JMP ww_AECreateList(SB)

TEXT ·AEPutDesc(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AEPutDesc(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AEPutDesc(SB),NOSPLIT,$0-0
 JMP ww_AEPutDesc(SB)

TEXT ·AECountItems(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AECountItems(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AECountItems(SB),NOSPLIT,$0-0
 JMP ww_AECountItems(SB)

TEXT ·AEGetNthDesc(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AEGetNthDesc(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AEGetNthDesc(SB),NOSPLIT,$0-0
 JMP ww_AEGetNthDesc(SB)

TEXT ·AEDuplicateDesc(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AEDuplicateDesc(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AEDuplicateDesc(SB),NOSPLIT,$0-0
 JMP ww_AEDuplicateDesc(SB)

TEXT ·CreateObjSpecifier(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_CreateObjSpecifier(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_CreateObjSpecifier(SB),NOSPLIT,$0-0
 JMP ww_CreateObjSpecifier(SB)

TEXT ·AECoerceDesc(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_AECoerceDesc(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_AECoerceDesc(SB),NOSPLIT,$0-0
 JMP ww_AECoerceDesc(SB)


TEXT ·ProcListAllPIDs(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_ProcListAllPIDs(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_ProcListAllPIDs(SB),NOSPLIT,$0-0
 JMP ww_ProcListAllPIDs(SB)

TEXT ·ProcPIDPath(SB),NOSPLIT,$0-8
 LEAQ ·trampoline_ProcPIDPath(SB), AX
 MOVQ AX, ret+0(FP)
 RET
TEXT ·trampoline_ProcPIDPath(SB),NOSPLIT,$0-0
 JMP ww_ProcPIDPath(SB)
