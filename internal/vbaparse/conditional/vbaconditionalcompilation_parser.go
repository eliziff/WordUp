// Code generated from VBAConditionalCompilationParser.g4 by ANTLR 4.13.1. DO NOT EDIT.

package conditional // VBAConditionalCompilationParser
import (
	"fmt"
	"strconv"
  	"sync"

	"github.com/antlr4-go/antlr/v4"
)

// Suppress unused import errors
var _ = fmt.Printf
var _ = strconv.Itoa
var _ = sync.Once{}


type VBAConditionalCompilationParser struct {
	*antlr.BaseParser
}

var VBAConditionalCompilationParserParserStaticData struct {
  once                   sync.Once
  serializedATN          []int32
  LiteralNames           []string
  SymbolicNames          []string
  RuleNames              []string
  PredictionContextCache *antlr.PredictionContextCache
  atn                    *antlr.ATN
  decisionToDFA          []*antlr.DFA
}

func vbaconditionalcompilationparserParserInit() {
  staticData := &VBAConditionalCompilationParserParserStaticData
  staticData.LiteralNames = []string{
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "','", "':'", "';'", "'!'", "'.'", "'#'", "'@'", "'%'", 
    "'$'", "'&'", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
    "", "", "", "", "", "", "", "", "", "", "':='", "'/'", "'\\'", "'='", 
    "", "'>'", "", "'('", "'<'", "'-'", "'*'", "", "'+'", "'^'", "')'", 
    "'['", "']'", "'{'", "'}'", "", "", "", "", "", "", "", "'''", "'_'",
  }
  staticData.SymbolicNames = []string{
    "", "ABS", "ANY", "ARRAY", "CBOOL", "CBYTE", "CCUR", "CDATE", "CDBL", 
    "CDEC", "CINT", "CIRCLE", "CLNG", "CLNGLNG", "CLNGPTR", "CSNG", "CSTR", 
    "CURRENCY", "CVAR", "CVERR", "DEBUG", "DOEVENTS", "EXIT", "FIX", "INPUTB", 
    "INT", "LBOUND", "LEN", "LENB", "LONGLONG", "LONGPTR", "MIDB", "OPTION", 
    "PSET", "SCALE", "SGN", "UBOUND", "COMMA", "COLON", "SEMICOLON", "EXCLAMATIONPOINT", 
    "DOT", "HASH", "AT", "PERCENT", "DOLLAR", "AMPERSAND", "ACCESS", "ADDRESSOF", 
    "ALIAS", "AND", "ATTRIBUTE", "APPEND", "AS", "BEGINPROPERTY", "BEGIN", 
    "BINARY", "BOOLEAN", "BYVAL", "BYREF", "BYTE", "CALL", "CASE", "CDECL", 
    "CLASS", "CLOSE", "CONST", "DATABASE", "DATE", "DECLARE", "DEFBOOL", 
    "DEFBYTE", "DEFDATE", "DEFDBL", "DEFCUR", "DEFINT", "DEFLNG", "DEFLNGLNG", 
    "DEFLNGPTR", "DEFOBJ", "DEFSNG", "DEFSTR", "DEFVAR", "DIM", "DO", "DOUBLE", 
    "EACH", "ELSE", "ELSEIF", "EMPTY", "END_ENUM", "END_FUNCTION", "END_IF", 
    "ENDPROPERTY", "END_PROPERTY", "END_SELECT", "END_SUB", "END_TYPE", 
    "END_WITH", "END", "ENUM", "EQV", "ERASE", "ERROR", "EVENT", "EXIT_DO", 
    "EXIT_FOR", "EXIT_FUNCTION", "EXIT_PROPERTY", "EXIT_SUB", "FALSE", "FRIEND", 
    "FOR", "FUNCTION", "GET", "GLOBAL", "GOSUB", "GOTO", "IF", "IMP", "IMPLEMENTS", 
    "IN", "INPUT", "IS", "INTEGER", "LOCK", "LONG", "LOOP", "LET", "LIB", 
    "LIKE", "LINE_INPUT", "LOCK_READ", "LOCK_WRITE", "LOCK_READ_WRITE", 
    "LSET", "ME", "MID", "MOD", "NAME", "NEXT", "NEW", "NOT", "NOTHING", 
    "NULL", "OBJECT", "ON", "ON_ERROR", "ON_LOCAL_ERROR", "OPEN", "OPTIONAL", 
    "OPTION_BASE", "OPTION_EXPLICIT", "OPTION_COMPARE", "OPTION_PRIVATE_MODULE", 
    "OR", "OUTPUT", "PARAMARRAY", "PRESERVE", "PRINT", "PRIVATE", "PROPERTY_GET", 
    "PROPERTY_LET", "PROPERTY_SET", "PTRSAFE", "PUBLIC", "PUT", "RANDOM", 
    "RANDOMIZE", "RAISEEVENT", "READ", "READ_WRITE", "REDIM", "REM", "RESET", 
    "RESUME", "RETURN", "RSET", "SEEK", "SELECT", "SET", "SHARED", "SINGLE", 
    "SPC", "STATIC", "STEP", "STOP", "STRING", "SUB", "TAB", "TEXT", "THEN", 
    "TO", "TRUE", "TYPE", "TYPEOF", "UNLOCK", "UNTIL", "VARIANT", "VERSION", 
    "WEND", "WHILE", "WIDTH", "WITH", "WITHEVENTS", "WRITE", "XOR", "ASSIGN", 
    "DIV", "INTDIV", "EQ", "GEQ", "GT", "LEQ", "LPAREN", "LT", "MINUS", 
    "MULT", "NEQ", "PLUS", "POW", "RPAREN", "L_SQUARE_BRACKET", "R_SQUARE_BRACKET", 
    "L_BRACE", "R_BRACE", "STRINGLITERAL", "OCTLITERAL", "HEXLITERAL", "FLOATLITERAL", 
    "INTEGERLITERAL", "DATELITERAL", "NEWLINE", "SINGLEQUOTE", "UNDERSCORE", 
    "WS", "GUIDLITERAL", "IDENTIFIER", "LINE_CONTINUATION", "BARE_HEX_LITERAL", 
    "ERRORCHAR", "LOAD", "MIDBTYPESUFFIX", "MIDTYPESUFFIX", "RESUME_NEXT",
  }
  staticData.RuleNames = []string{
    "compilationUnit", "ccBlock", "ccConst", "ccVarLhs", "ccIfBlock", "ccIf", 
    "ccElseIfBlock", "ccElseIf", "ccElseBlock", "ccElse", "ccEndIf", "ccEol", 
    "hashConst", "hashIf", "hashElseIf", "hashElse", "hashEndIf", "physicalLine", 
    "ccExpression", "intrinsicFunction", "intrinsicFunctionName", "name", 
    "nameValue", "foreignName", "foreignIdentifier", "typeHint", "literal", 
    "comment", "keyword", "markerKeyword", "statementKeyword", "whiteSpace",
  }
  staticData.PredictionContextCache = antlr.NewPredictionContextCache()
  staticData.serializedATN = []int32{
	4, 1, 244, 509, 2, 0, 7, 0, 2, 1, 7, 1, 2, 2, 7, 2, 2, 3, 7, 3, 2, 4, 7, 
	4, 2, 5, 7, 5, 2, 6, 7, 6, 2, 7, 7, 7, 2, 8, 7, 8, 2, 9, 7, 9, 2, 10, 7, 
	10, 2, 11, 7, 11, 2, 12, 7, 12, 2, 13, 7, 13, 2, 14, 7, 14, 2, 15, 7, 15, 
	2, 16, 7, 16, 2, 17, 7, 17, 2, 18, 7, 18, 2, 19, 7, 19, 2, 20, 7, 20, 2, 
	21, 7, 21, 2, 22, 7, 22, 2, 23, 7, 23, 2, 24, 7, 24, 2, 25, 7, 25, 2, 26, 
	7, 26, 2, 27, 7, 27, 2, 28, 7, 28, 2, 29, 7, 29, 2, 30, 7, 30, 2, 31, 7, 
	31, 1, 0, 1, 0, 1, 0, 1, 1, 1, 1, 1, 1, 5, 1, 71, 8, 1, 10, 1, 12, 1, 74, 
	9, 1, 1, 2, 5, 2, 77, 8, 2, 10, 2, 12, 2, 80, 9, 2, 1, 2, 1, 2, 4, 2, 84, 
	8, 2, 11, 2, 12, 2, 85, 1, 2, 1, 2, 5, 2, 90, 8, 2, 10, 2, 12, 2, 93, 9, 
	2, 1, 2, 1, 2, 5, 2, 97, 8, 2, 10, 2, 12, 2, 100, 9, 2, 1, 2, 1, 2, 1, 
	2, 1, 3, 1, 3, 1, 4, 1, 4, 1, 4, 5, 4, 110, 8, 4, 10, 4, 12, 4, 113, 9, 
	4, 1, 4, 3, 4, 116, 8, 4, 1, 4, 1, 4, 1, 5, 5, 5, 121, 8, 5, 10, 5, 12, 
	5, 124, 9, 5, 1, 5, 1, 5, 4, 5, 128, 8, 5, 11, 5, 12, 5, 129, 1, 5, 1, 
	5, 4, 5, 134, 8, 5, 11, 5, 12, 5, 135, 1, 5, 1, 5, 1, 5, 1, 6, 1, 6, 1, 
	6, 1, 7, 5, 7, 145, 8, 7, 10, 7, 12, 7, 148, 9, 7, 1, 7, 1, 7, 4, 7, 152, 
	8, 7, 11, 7, 12, 7, 153, 1, 7, 1, 7, 4, 7, 158, 8, 7, 11, 7, 12, 7, 159, 
	1, 7, 1, 7, 1, 7, 1, 8, 1, 8, 1, 8, 1, 9, 5, 9, 169, 8, 9, 10, 9, 12, 9, 
	172, 9, 9, 1, 9, 1, 9, 1, 9, 1, 10, 5, 10, 178, 8, 10, 10, 10, 12, 10, 
	181, 9, 10, 1, 10, 1, 10, 1, 10, 1, 11, 5, 11, 187, 8, 11, 10, 11, 12, 
	11, 190, 9, 11, 1, 11, 3, 11, 193, 8, 11, 1, 11, 1, 11, 1, 12, 1, 12, 1, 
	12, 1, 13, 1, 13, 1, 13, 1, 14, 1, 14, 1, 14, 1, 15, 1, 15, 1, 15, 1, 16, 
	1, 16, 1, 16, 1, 17, 5, 17, 213, 8, 17, 10, 17, 12, 17, 216, 9, 17, 1, 
	17, 1, 17, 1, 18, 1, 18, 1, 18, 5, 18, 223, 8, 18, 10, 18, 12, 18, 226, 
	9, 18, 1, 18, 1, 18, 5, 18, 230, 8, 18, 10, 18, 12, 18, 233, 9, 18, 1, 
	18, 1, 18, 1, 18, 1, 18, 5, 18, 239, 8, 18, 10, 18, 12, 18, 242, 9, 18, 
	1, 18, 1, 18, 1, 18, 5, 18, 247, 8, 18, 10, 18, 12, 18, 250, 9, 18, 1, 
	18, 1, 18, 1, 18, 1, 18, 3, 18, 256, 8, 18, 1, 18, 1, 18, 5, 18, 260, 8, 
	18, 10, 18, 12, 18, 263, 9, 18, 1, 18, 1, 18, 5, 18, 267, 8, 18, 10, 18, 
	12, 18, 270, 9, 18, 1, 18, 1, 18, 1, 18, 5, 18, 275, 8, 18, 10, 18, 12, 
	18, 278, 9, 18, 1, 18, 1, 18, 5, 18, 282, 8, 18, 10, 18, 12, 18, 285, 9, 
	18, 1, 18, 1, 18, 1, 18, 5, 18, 290, 8, 18, 10, 18, 12, 18, 293, 9, 18, 
	1, 18, 1, 18, 5, 18, 297, 8, 18, 10, 18, 12, 18, 300, 9, 18, 1, 18, 1, 
	18, 1, 18, 5, 18, 305, 8, 18, 10, 18, 12, 18, 308, 9, 18, 1, 18, 1, 18, 
	5, 18, 312, 8, 18, 10, 18, 12, 18, 315, 9, 18, 1, 18, 1, 18, 1, 18, 5, 
	18, 320, 8, 18, 10, 18, 12, 18, 323, 9, 18, 1, 18, 1, 18, 5, 18, 327, 8, 
	18, 10, 18, 12, 18, 330, 9, 18, 1, 18, 1, 18, 1, 18, 5, 18, 335, 8, 18, 
	10, 18, 12, 18, 338, 9, 18, 1, 18, 1, 18, 5, 18, 342, 8, 18, 10, 18, 12, 
	18, 345, 9, 18, 1, 18, 1, 18, 1, 18, 5, 18, 350, 8, 18, 10, 18, 12, 18, 
	353, 9, 18, 1, 18, 1, 18, 5, 18, 357, 8, 18, 10, 18, 12, 18, 360, 9, 18, 
	1, 18, 1, 18, 1, 18, 5, 18, 365, 8, 18, 10, 18, 12, 18, 368, 9, 18, 1, 
	18, 1, 18, 5, 18, 372, 8, 18, 10, 18, 12, 18, 375, 9, 18, 1, 18, 1, 18, 
	1, 18, 5, 18, 380, 8, 18, 10, 18, 12, 18, 383, 9, 18, 1, 18, 1, 18, 5, 
	18, 387, 8, 18, 10, 18, 12, 18, 390, 9, 18, 1, 18, 1, 18, 1, 18, 5, 18, 
	395, 8, 18, 10, 18, 12, 18, 398, 9, 18, 1, 18, 1, 18, 5, 18, 402, 8, 18, 
	10, 18, 12, 18, 405, 9, 18, 1, 18, 1, 18, 1, 18, 5, 18, 410, 8, 18, 10, 
	18, 12, 18, 413, 9, 18, 1, 18, 1, 18, 5, 18, 417, 8, 18, 10, 18, 12, 18, 
	420, 9, 18, 1, 18, 1, 18, 1, 18, 5, 18, 425, 8, 18, 10, 18, 12, 18, 428, 
	9, 18, 1, 18, 1, 18, 5, 18, 432, 8, 18, 10, 18, 12, 18, 435, 9, 18, 1, 
	18, 5, 18, 438, 8, 18, 10, 18, 12, 18, 441, 9, 18, 1, 19, 1, 19, 1, 19, 
	5, 19, 446, 8, 19, 10, 19, 12, 19, 449, 9, 19, 1, 19, 1, 19, 5, 19, 453, 
	8, 19, 10, 19, 12, 19, 456, 9, 19, 1, 19, 1, 19, 1, 20, 1, 20, 1, 21, 1, 
	21, 3, 21, 464, 8, 21, 1, 22, 1, 22, 1, 22, 1, 22, 1, 22, 3, 22, 471, 8, 
	22, 1, 23, 1, 23, 5, 23, 475, 8, 23, 10, 23, 12, 23, 478, 9, 23, 1, 23, 
	1, 23, 1, 24, 1, 24, 3, 24, 484, 8, 24, 1, 25, 1, 25, 1, 26, 1, 26, 1, 
	27, 1, 27, 1, 27, 5, 27, 493, 8, 27, 10, 27, 12, 27, 496, 9, 27, 1, 28, 
	1, 28, 1, 29, 1, 29, 1, 30, 1, 30, 1, 31, 4, 31, 505, 8, 31, 11, 31, 12, 
	31, 506, 1, 31, 1, 72, 1, 36, 32, 0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 
	22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56, 
	58, 60, 62, 0, 11, 1, 0, 232, 232, 2, 0, 208, 208, 217, 217, 2, 0, 216, 
	216, 219, 219, 5, 0, 123, 123, 130, 130, 210, 213, 215, 215, 218, 218, 
	9, 0, 1, 1, 4, 8, 10, 10, 12, 16, 18, 18, 23, 23, 25, 25, 27, 28, 35, 35, 
	1, 0, 222, 222, 3, 0, 40, 40, 42, 46, 220, 220, 5, 0, 89, 89, 110, 110, 
	143, 144, 193, 193, 226, 231, 36, 0, 1, 10, 12, 21, 23, 31, 33, 33, 35, 
	36, 47, 52, 55, 60, 64, 65, 67, 68, 85, 85, 99, 99, 101, 101, 103, 103, 
	110, 110, 114, 114, 119, 119, 121, 126, 129, 134, 136, 139, 141, 145, 147, 
	147, 149, 150, 155, 159, 164, 164, 166, 167, 170, 171, 173, 174, 178, 178, 
	181, 183, 185, 185, 187, 187, 189, 193, 195, 199, 202, 202, 204, 206, 241, 
	244, 30, 0, 22, 22, 32, 32, 61, 62, 66, 66, 69, 84, 87, 88, 95, 95, 98, 
	98, 100, 100, 102, 102, 104, 109, 111, 113, 115, 118, 120, 120, 127, 128, 
	135, 135, 140, 140, 146, 146, 160, 160, 165, 165, 169, 169, 172, 172, 175, 
	177, 179, 180, 184, 184, 186, 186, 188, 188, 194, 194, 200, 201, 203, 203, 
	2, 0, 235, 235, 238, 238, 553, 0, 64, 1, 0, 0, 0, 2, 72, 1, 0, 0, 0, 4, 
	78, 1, 0, 0, 0, 6, 104, 1, 0, 0, 0, 8, 106, 1, 0, 0, 0, 10, 122, 1, 0, 
	0, 0, 12, 140, 1, 0, 0, 0, 14, 146, 1, 0, 0, 0, 16, 164, 1, 0, 0, 0, 18, 
	170, 1, 0, 0, 0, 20, 179, 1, 0, 0, 0, 22, 188, 1, 0, 0, 0, 24, 196, 1, 
	0, 0, 0, 26, 199, 1, 0, 0, 0, 28, 202, 1, 0, 0, 0, 30, 205, 1, 0, 0, 0, 
	32, 208, 1, 0, 0, 0, 34, 214, 1, 0, 0, 0, 36, 255, 1, 0, 0, 0, 38, 442, 
	1, 0, 0, 0, 40, 459, 1, 0, 0, 0, 42, 461, 1, 0, 0, 0, 44, 470, 1, 0, 0, 
	0, 46, 472, 1, 0, 0, 0, 48, 483, 1, 0, 0, 0, 50, 485, 1, 0, 0, 0, 52, 487, 
	1, 0, 0, 0, 54, 489, 1, 0, 0, 0, 56, 497, 1, 0, 0, 0, 58, 499, 1, 0, 0, 
	0, 60, 501, 1, 0, 0, 0, 62, 504, 1, 0, 0, 0, 64, 65, 3, 2, 1, 0, 65, 66, 
	5, 0, 0, 1, 66, 1, 1, 0, 0, 0, 67, 71, 3, 4, 2, 0, 68, 71, 3, 8, 4, 0, 
	69, 71, 3, 34, 17, 0, 70, 67, 1, 0, 0, 0, 70, 68, 1, 0, 0, 0, 70, 69, 1, 
	0, 0, 0, 71, 74, 1, 0, 0, 0, 72, 73, 1, 0, 0, 0, 72, 70, 1, 0, 0, 0, 73, 
	3, 1, 0, 0, 0, 74, 72, 1, 0, 0, 0, 75, 77, 3, 62, 31, 0, 76, 75, 1, 0, 
	0, 0, 77, 80, 1, 0, 0, 0, 78, 76, 1, 0, 0, 0, 78, 79, 1, 0, 0, 0, 79, 81, 
	1, 0, 0, 0, 80, 78, 1, 0, 0, 0, 81, 83, 3, 24, 12, 0, 82, 84, 3, 62, 31, 
	0, 83, 82, 1, 0, 0, 0, 84, 85, 1, 0, 0, 0, 85, 83, 1, 0, 0, 0, 85, 86, 
	1, 0, 0, 0, 86, 87, 1, 0, 0, 0, 87, 91, 3, 6, 3, 0, 88, 90, 3, 62, 31, 
	0, 89, 88, 1, 0, 0, 0, 90, 93, 1, 0, 0, 0, 91, 89, 1, 0, 0, 0, 91, 92, 
	1, 0, 0, 0, 92, 94, 1, 0, 0, 0, 93, 91, 1, 0, 0, 0, 94, 98, 5, 210, 0, 
	0, 95, 97, 3, 62, 31, 0, 96, 95, 1, 0, 0, 0, 97, 100, 1, 0, 0, 0, 98, 96, 
	1, 0, 0, 0, 98, 99, 1, 0, 0, 0, 99, 101, 1, 0, 0, 0, 100, 98, 1, 0, 0, 
	0, 101, 102, 3, 36, 18, 0, 102, 103, 3, 22, 11, 0, 103, 5, 1, 0, 0, 0, 
	104, 105, 3, 42, 21, 0, 105, 7, 1, 0, 0, 0, 106, 107, 3, 10, 5, 0, 107, 
	111, 3, 2, 1, 0, 108, 110, 3, 12, 6, 0, 109, 108, 1, 0, 0, 0, 110, 113, 
	1, 0, 0, 0, 111, 109, 1, 0, 0, 0, 111, 112, 1, 0, 0, 0, 112, 115, 1, 0, 
	0, 0, 113, 111, 1, 0, 0, 0, 114, 116, 3, 16, 8, 0, 115, 114, 1, 0, 0, 0, 
	115, 116, 1, 0, 0, 0, 116, 117, 1, 0, 0, 0, 117, 118, 3, 20, 10, 0, 118, 
	9, 1, 0, 0, 0, 119, 121, 3, 62, 31, 0, 120, 119, 1, 0, 0, 0, 121, 124, 
	1, 0, 0, 0, 122, 120, 1, 0, 0, 0, 122, 123, 1, 0, 0, 0, 123, 125, 1, 0, 
	0, 0, 124, 122, 1, 0, 0, 0, 125, 127, 3, 26, 13, 0, 126, 128, 3, 62, 31, 
	0, 127, 126, 1, 0, 0, 0, 128, 129, 1, 0, 0, 0, 129, 127, 1, 0, 0, 0, 129, 
	130, 1, 0, 0, 0, 130, 131, 1, 0, 0, 0, 131, 133, 3, 36, 18, 0, 132, 134, 
	3, 62, 31, 0, 133, 132, 1, 0, 0, 0, 134, 135, 1, 0, 0, 0, 135, 133, 1, 
	0, 0, 0, 135, 136, 1, 0, 0, 0, 136, 137, 1, 0, 0, 0, 137, 138, 5, 191, 
	0, 0, 138, 139, 3, 22, 11, 0, 139, 11, 1, 0, 0, 0, 140, 141, 3, 14, 7, 
	0, 141, 142, 3, 2, 1, 0, 142, 13, 1, 0, 0, 0, 143, 145, 3, 62, 31, 0, 144, 
	143, 1, 0, 0, 0, 145, 148, 1, 0, 0, 0, 146, 144, 1, 0, 0, 0, 146, 147, 
	1, 0, 0, 0, 147, 149, 1, 0, 0, 0, 148, 146, 1, 0, 0, 0, 149, 151, 3, 28, 
	14, 0, 150, 152, 3, 62, 31, 0, 151, 150, 1, 0, 0, 0, 152, 153, 1, 0, 0, 
	0, 153, 151, 1, 0, 0, 0, 153, 154, 1, 0, 0, 0, 154, 155, 1, 0, 0, 0, 155, 
	157, 3, 36, 18, 0, 156, 158, 3, 62, 31, 0, 157, 156, 1, 0, 0, 0, 158, 159, 
	1, 0, 0, 0, 159, 157, 1, 0, 0, 0, 159, 160, 1, 0, 0, 0, 160, 161, 1, 0, 
	0, 0, 161, 162, 5, 191, 0, 0, 162, 163, 3, 22, 11, 0, 163, 15, 1, 0, 0, 
	0, 164, 165, 3, 18, 9, 0, 165, 166, 3, 2, 1, 0, 166, 17, 1, 0, 0, 0, 167, 
	169, 3, 62, 31, 0, 168, 167, 1, 0, 0, 0, 169, 172, 1, 0, 0, 0, 170, 168, 
	1, 0, 0, 0, 170, 171, 1, 0, 0, 0, 171, 173, 1, 0, 0, 0, 172, 170, 1, 0, 
	0, 0, 173, 174, 3, 30, 15, 0, 174, 175, 3, 22, 11, 0, 175, 19, 1, 0, 0, 
	0, 176, 178, 3, 62, 31, 0, 177, 176, 1, 0, 0, 0, 178, 181, 1, 0, 0, 0, 
	179, 177, 1, 0, 0, 0, 179, 180, 1, 0, 0, 0, 180, 182, 1, 0, 0, 0, 181, 
	179, 1, 0, 0, 0, 182, 183, 3, 32, 16, 0, 183, 184, 3, 22, 11, 0, 184, 21, 
	1, 0, 0, 0, 185, 187, 3, 62, 31, 0, 186, 185, 1, 0, 0, 0, 187, 190, 1, 
	0, 0, 0, 188, 186, 1, 0, 0, 0, 188, 189, 1, 0, 0, 0, 189, 192, 1, 0, 0, 
	0, 190, 188, 1, 0, 0, 0, 191, 193, 3, 54, 27, 0, 192, 191, 1, 0, 0, 0, 
	192, 193, 1, 0, 0, 0, 193, 194, 1, 0, 0, 0, 194, 195, 5, 232, 0, 0, 195, 
	23, 1, 0, 0, 0, 196, 197, 5, 42, 0, 0, 197, 198, 5, 66, 0, 0, 198, 25, 
	1, 0, 0, 0, 199, 200, 5, 42, 0, 0, 200, 201, 5, 118, 0, 0, 201, 27, 1, 
	0, 0, 0, 202, 203, 5, 42, 0, 0, 203, 204, 5, 88, 0, 0, 204, 29, 1, 0, 0, 
	0, 205, 206, 5, 42, 0, 0, 206, 207, 5, 87, 0, 0, 207, 31, 1, 0, 0, 0, 208, 
	209, 5, 42, 0, 0, 209, 210, 5, 92, 0, 0, 210, 33, 1, 0, 0, 0, 211, 213, 
	8, 0, 0, 0, 212, 211, 1, 0, 0, 0, 213, 216, 1, 0, 0, 0, 214, 212, 1, 0, 
	0, 0, 214, 215, 1, 0, 0, 0, 215, 217, 1, 0, 0, 0, 216, 214, 1, 0, 0, 0, 
	217, 218, 5, 232, 0, 0, 218, 35, 1, 0, 0, 0, 219, 220, 6, 18, -1, 0, 220, 
	224, 5, 214, 0, 0, 221, 223, 3, 62, 31, 0, 222, 221, 1, 0, 0, 0, 223, 226, 
	1, 0, 0, 0, 224, 222, 1, 0, 0, 0, 224, 225, 1, 0, 0, 0, 225, 227, 1, 0, 
	0, 0, 226, 224, 1, 0, 0, 0, 227, 231, 3, 36, 18, 0, 228, 230, 3, 62, 31, 
	0, 229, 228, 1, 0, 0, 0, 230, 233, 1, 0, 0, 0, 231, 229, 1, 0, 0, 0, 231, 
	232, 1, 0, 0, 0, 232, 234, 1, 0, 0, 0, 233, 231, 1, 0, 0, 0, 234, 235, 
	5, 221, 0, 0, 235, 256, 1, 0, 0, 0, 236, 240, 5, 216, 0, 0, 237, 239, 3, 
	62, 31, 0, 238, 237, 1, 0, 0, 0, 239, 242, 1, 0, 0, 0, 240, 238, 1, 0, 
	0, 0, 240, 241, 1, 0, 0, 0, 241, 243, 1, 0, 0, 0, 242, 240, 1, 0, 0, 0, 
	243, 256, 3, 36, 18, 16, 244, 248, 5, 142, 0, 0, 245, 247, 3, 62, 31, 0, 
	246, 245, 1, 0, 0, 0, 247, 250, 1, 0, 0, 0, 248, 246, 1, 0, 0, 0, 248, 
	249, 1, 0, 0, 0, 249, 251, 1, 0, 0, 0, 250, 248, 1, 0, 0, 0, 251, 256, 
	3, 36, 18, 9, 252, 256, 3, 38, 19, 0, 253, 256, 3, 52, 26, 0, 254, 256, 
	3, 42, 21, 0, 255, 219, 1, 0, 0, 0, 255, 236, 1, 0, 0, 0, 255, 244, 1, 
	0, 0, 0, 255, 252, 1, 0, 0, 0, 255, 253, 1, 0, 0, 0, 255, 254, 1, 0, 0, 
	0, 256, 439, 1, 0, 0, 0, 257, 261, 10, 17, 0, 0, 258, 260, 3, 62, 31, 0, 
	259, 258, 1, 0, 0, 0, 260, 263, 1, 0, 0, 0, 261, 259, 1, 0, 0, 0, 261, 
	262, 1, 0, 0, 0, 262, 264, 1, 0, 0, 0, 263, 261, 1, 0, 0, 0, 264, 268, 
	5, 220, 0, 0, 265, 267, 3, 62, 31, 0, 266, 265, 1, 0, 0, 0, 267, 270, 1, 
	0, 0, 0, 268, 266, 1, 0, 0, 0, 268, 269, 1, 0, 0, 0, 269, 271, 1, 0, 0, 
	0, 270, 268, 1, 0, 0, 0, 271, 438, 3, 36, 18, 18, 272, 276, 10, 15, 0, 
	0, 273, 275, 3, 62, 31, 0, 274, 273, 1, 0, 0, 0, 275, 278, 1, 0, 0, 0, 
	276, 274, 1, 0, 0, 0, 276, 277, 1, 0, 0, 0, 277, 279, 1, 0, 0, 0, 278, 
	276, 1, 0, 0, 0, 279, 283, 7, 1, 0, 0, 280, 282, 3, 62, 31, 0, 281, 280, 
	1, 0, 0, 0, 282, 285, 1, 0, 0, 0, 283, 281, 1, 0, 0, 0, 283, 284, 1, 0, 
	0, 0, 284, 286, 1, 0, 0, 0, 285, 283, 1, 0, 0, 0, 286, 438, 3, 36, 18, 
	16, 287, 291, 10, 14, 0, 0, 288, 290, 3, 62, 31, 0, 289, 288, 1, 0, 0, 
	0, 290, 293, 1, 0, 0, 0, 291, 289, 1, 0, 0, 0, 291, 292, 1, 0, 0, 0, 292, 
	294, 1, 0, 0, 0, 293, 291, 1, 0, 0, 0, 294, 298, 5, 209, 0, 0, 295, 297, 
	3, 62, 31, 0, 296, 295, 1, 0, 0, 0, 297, 300, 1, 0, 0, 0, 298, 296, 1, 
	0, 0, 0, 298, 299, 1, 0, 0, 0, 299, 301, 1, 0, 0, 0, 300, 298, 1, 0, 0, 
	0, 301, 438, 3, 36, 18, 15, 302, 306, 10, 13, 0, 0, 303, 305, 3, 62, 31, 
	0, 304, 303, 1, 0, 0, 0, 305, 308, 1, 0, 0, 0, 306, 304, 1, 0, 0, 0, 306, 
	307, 1, 0, 0, 0, 307, 309, 1, 0, 0, 0, 308, 306, 1, 0, 0, 0, 309, 313, 
	5, 138, 0, 0, 310, 312, 3, 62, 31, 0, 311, 310, 1, 0, 0, 0, 312, 315, 1, 
	0, 0, 0, 313, 311, 1, 0, 0, 0, 313, 314, 1, 0, 0, 0, 314, 316, 1, 0, 0, 
	0, 315, 313, 1, 0, 0, 0, 316, 438, 3, 36, 18, 14, 317, 321, 10, 12, 0, 
	0, 318, 320, 3, 62, 31, 0, 319, 318, 1, 0, 0, 0, 320, 323, 1, 0, 0, 0, 
	321, 319, 1, 0, 0, 0, 321, 322, 1, 0, 0, 0, 322, 324, 1, 0, 0, 0, 323, 
	321, 1, 0, 0, 0, 324, 328, 7, 2, 0, 0, 325, 327, 3, 62, 31, 0, 326, 325, 
	1, 0, 0, 0, 327, 330, 1, 0, 0, 0, 328, 326, 1, 0, 0, 0, 328, 329, 1, 0, 
	0, 0, 329, 331, 1, 0, 0, 0, 330, 328, 1, 0, 0, 0, 331, 438, 3, 36, 18, 
	13, 332, 336, 10, 11, 0, 0, 333, 335, 3, 62, 31, 0, 334, 333, 1, 0, 0, 
	0, 335, 338, 1, 0, 0, 0, 336, 334, 1, 0, 0, 0, 336, 337, 1, 0, 0, 0, 337, 
	339, 1, 0, 0, 0, 338, 336, 1, 0, 0, 0, 339, 343, 5, 46, 0, 0, 340, 342, 
	3, 62, 31, 0, 341, 340, 1, 0, 0, 0, 342, 345, 1, 0, 0, 0, 343, 341, 1, 
	0, 0, 0, 343, 344, 1, 0, 0, 0, 344, 346, 1, 0, 0, 0, 345, 343, 1, 0, 0, 
	0, 346, 438, 3, 36, 18, 12, 347, 351, 10, 10, 0, 0, 348, 350, 3, 62, 31, 
	0, 349, 348, 1, 0, 0, 0, 350, 353, 1, 0, 0, 0, 351, 349, 1, 0, 0, 0, 351, 
	352, 1, 0, 0, 0, 352, 354, 1, 0, 0, 0, 353, 351, 1, 0, 0, 0, 354, 358, 
	7, 3, 0, 0, 355, 357, 3, 62, 31, 0, 356, 355, 1, 0, 0, 0, 357, 360, 1, 
	0, 0, 0, 358, 356, 1, 0, 0, 0, 358, 359, 1, 0, 0, 0, 359, 361, 1, 0, 0, 
	0, 360, 358, 1, 0, 0, 0, 361, 438, 3, 36, 18, 11, 362, 366, 10, 8, 0, 0, 
	363, 365, 3, 62, 31, 0, 364, 363, 1, 0, 0, 0, 365, 368, 1, 0, 0, 0, 366, 
	364, 1, 0, 0, 0, 366, 367, 1, 0, 0, 0, 367, 369, 1, 0, 0, 0, 368, 366, 
	1, 0, 0, 0, 369, 373, 5, 50, 0, 0, 370, 372, 3, 62, 31, 0, 371, 370, 1, 
	0, 0, 0, 372, 375, 1, 0, 0, 0, 373, 371, 1, 0, 0, 0, 373, 374, 1, 0, 0, 
	0, 374, 376, 1, 0, 0, 0, 375, 373, 1, 0, 0, 0, 376, 438, 3, 36, 18, 9, 
	377, 381, 10, 7, 0, 0, 378, 380, 3, 62, 31, 0, 379, 378, 1, 0, 0, 0, 380, 
	383, 1, 0, 0, 0, 381, 379, 1, 0, 0, 0, 381, 382, 1, 0, 0, 0, 382, 384, 
	1, 0, 0, 0, 383, 381, 1, 0, 0, 0, 384, 388, 5, 155, 0, 0, 385, 387, 3, 
	62, 31, 0, 386, 385, 1, 0, 0, 0, 387, 390, 1, 0, 0, 0, 388, 386, 1, 0, 
	0, 0, 388, 389, 1, 0, 0, 0, 389, 391, 1, 0, 0, 0, 390, 388, 1, 0, 0, 0, 
	391, 438, 3, 36, 18, 8, 392, 396, 10, 6, 0, 0, 393, 395, 3, 62, 31, 0, 
	394, 393, 1, 0, 0, 0, 395, 398, 1, 0, 0, 0, 396, 394, 1, 0, 0, 0, 396, 
	397, 1, 0, 0, 0, 397, 399, 1, 0, 0, 0, 398, 396, 1, 0, 0, 0, 399, 403, 
	5, 206, 0, 0, 400, 402, 3, 62, 31, 0, 401, 400, 1, 0, 0, 0, 402, 405, 1, 
	0, 0, 0, 403, 401, 1, 0, 0, 0, 403, 404, 1, 0, 0, 0, 404, 406, 1, 0, 0, 
	0, 405, 403, 1, 0, 0, 0, 406, 438, 3, 36, 18, 7, 407, 411, 10, 5, 0, 0, 
	408, 410, 3, 62, 31, 0, 409, 408, 1, 0, 0, 0, 410, 413, 1, 0, 0, 0, 411, 
	409, 1, 0, 0, 0, 411, 412, 1, 0, 0, 0, 412, 414, 1, 0, 0, 0, 413, 411, 
	1, 0, 0, 0, 414, 418, 5, 101, 0, 0, 415, 417, 3, 62, 31, 0, 416, 415, 1, 
	0, 0, 0, 417, 420, 1, 0, 0, 0, 418, 416, 1, 0, 0, 0, 418, 419, 1, 0, 0, 
	0, 419, 421, 1, 0, 0, 0, 420, 418, 1, 0, 0, 0, 421, 438, 3, 36, 18, 6, 
	422, 426, 10, 4, 0, 0, 423, 425, 3, 62, 31, 0, 424, 423, 1, 0, 0, 0, 425, 
	428, 1, 0, 0, 0, 426, 424, 1, 0, 0, 0, 426, 427, 1, 0, 0, 0, 427, 429, 
	1, 0, 0, 0, 428, 426, 1, 0, 0, 0, 429, 433, 5, 119, 0, 0, 430, 432, 3, 
	62, 31, 0, 431, 430, 1, 0, 0, 0, 432, 435, 1, 0, 0, 0, 433, 431, 1, 0, 
	0, 0, 433, 434, 1, 0, 0, 0, 434, 436, 1, 0, 0, 0, 435, 433, 1, 0, 0, 0, 
	436, 438, 3, 36, 18, 5, 437, 257, 1, 0, 0, 0, 437, 272, 1, 0, 0, 0, 437, 
	287, 1, 0, 0, 0, 437, 302, 1, 0, 0, 0, 437, 317, 1, 0, 0, 0, 437, 332, 
	1, 0, 0, 0, 437, 347, 1, 0, 0, 0, 437, 362, 1, 0, 0, 0, 437, 377, 1, 0, 
	0, 0, 437, 392, 1, 0, 0, 0, 437, 407, 1, 0, 0, 0, 437, 422, 1, 0, 0, 0, 
	438, 441, 1, 0, 0, 0, 439, 437, 1, 0, 0, 0, 439, 440, 1, 0, 0, 0, 440, 
	37, 1, 0, 0, 0, 441, 439, 1, 0, 0, 0, 442, 443, 3, 40, 20, 0, 443, 447, 
	5, 214, 0, 0, 444, 446, 3, 62, 31, 0, 445, 444, 1, 0, 0, 0, 446, 449, 1, 
	0, 0, 0, 447, 445, 1, 0, 0, 0, 447, 448, 1, 0, 0, 0, 448, 450, 1, 0, 0, 
	0, 449, 447, 1, 0, 0, 0, 450, 454, 3, 36, 18, 0, 451, 453, 3, 62, 31, 0, 
	452, 451, 1, 0, 0, 0, 453, 456, 1, 0, 0, 0, 454, 452, 1, 0, 0, 0, 454, 
	455, 1, 0, 0, 0, 455, 457, 1, 0, 0, 0, 456, 454, 1, 0, 0, 0, 457, 458, 
	5, 221, 0, 0, 458, 39, 1, 0, 0, 0, 459, 460, 7, 4, 0, 0, 460, 41, 1, 0, 
	0, 0, 461, 463, 3, 44, 22, 0, 462, 464, 3, 50, 25, 0, 463, 462, 1, 0, 0, 
	0, 463, 464, 1, 0, 0, 0, 464, 43, 1, 0, 0, 0, 465, 471, 5, 237, 0, 0, 466, 
	471, 3, 56, 28, 0, 467, 471, 3, 46, 23, 0, 468, 471, 3, 60, 30, 0, 469, 
	471, 3, 58, 29, 0, 470, 465, 1, 0, 0, 0, 470, 466, 1, 0, 0, 0, 470, 467, 
	1, 0, 0, 0, 470, 468, 1, 0, 0, 0, 470, 469, 1, 0, 0, 0, 471, 45, 1, 0, 
	0, 0, 472, 476, 5, 222, 0, 0, 473, 475, 3, 48, 24, 0, 474, 473, 1, 0, 0, 
	0, 475, 478, 1, 0, 0, 0, 476, 474, 1, 0, 0, 0, 476, 477, 1, 0, 0, 0, 477, 
	479, 1, 0, 0, 0, 478, 476, 1, 0, 0, 0, 479, 480, 5, 223, 0, 0, 480, 47, 
	1, 0, 0, 0, 481, 484, 8, 5, 0, 0, 482, 484, 3, 46, 23, 0, 483, 481, 1, 
	0, 0, 0, 483, 482, 1, 0, 0, 0, 484, 49, 1, 0, 0, 0, 485, 486, 7, 6, 0, 
	0, 486, 51, 1, 0, 0, 0, 487, 488, 7, 7, 0, 0, 488, 53, 1, 0, 0, 0, 489, 
	494, 5, 233, 0, 0, 490, 493, 5, 238, 0, 0, 491, 493, 8, 0, 0, 0, 492, 490, 
	1, 0, 0, 0, 492, 491, 1, 0, 0, 0, 493, 496, 1, 0, 0, 0, 494, 492, 1, 0, 
	0, 0, 494, 495, 1, 0, 0, 0, 495, 55, 1, 0, 0, 0, 496, 494, 1, 0, 0, 0, 
	497, 498, 7, 8, 0, 0, 498, 57, 1, 0, 0, 0, 499, 500, 5, 53, 0, 0, 500, 
	59, 1, 0, 0, 0, 501, 502, 7, 9, 0, 0, 502, 61, 1, 0, 0, 0, 503, 505, 7, 
	10, 0, 0, 504, 503, 1, 0, 0, 0, 505, 506, 1, 0, 0, 0, 506, 504, 1, 0, 0, 
	0, 506, 507, 1, 0, 0, 0, 507, 63, 1, 0, 0, 0, 59, 70, 72, 78, 85, 91, 98, 
	111, 115, 122, 129, 135, 146, 153, 159, 170, 179, 188, 192, 214, 224, 231, 
	240, 248, 255, 261, 268, 276, 283, 291, 298, 306, 313, 321, 328, 336, 343, 
	351, 358, 366, 373, 381, 388, 396, 403, 411, 418, 426, 433, 437, 439, 447, 
	454, 463, 470, 476, 483, 492, 494, 506,
}
  deserializer := antlr.NewATNDeserializer(nil)
  staticData.atn = deserializer.Deserialize(staticData.serializedATN)
  atn := staticData.atn
  staticData.decisionToDFA = make([]*antlr.DFA, len(atn.DecisionToState))
  decisionToDFA := staticData.decisionToDFA
  for index, state := range atn.DecisionToState {
    decisionToDFA[index] = antlr.NewDFA(state, index)
  }
}

// VBAConditionalCompilationParserInit initializes any static state used to implement VBAConditionalCompilationParser. By default the
// static state used to implement the parser is lazily initialized during the first call to
// NewVBAConditionalCompilationParser(). You can call this function if you wish to initialize the static state ahead
// of time.
func VBAConditionalCompilationParserInit() {
  staticData := &VBAConditionalCompilationParserParserStaticData
  staticData.once.Do(vbaconditionalcompilationparserParserInit)
}

// NewVBAConditionalCompilationParser produces a new parser instance for the optional input antlr.TokenStream.
func NewVBAConditionalCompilationParser(input antlr.TokenStream) *VBAConditionalCompilationParser {
	VBAConditionalCompilationParserInit()
	this := new(VBAConditionalCompilationParser)
	this.BaseParser = antlr.NewBaseParser(input)
  staticData := &VBAConditionalCompilationParserParserStaticData
	this.Interpreter = antlr.NewParserATNSimulator(this, staticData.atn, staticData.decisionToDFA, staticData.PredictionContextCache)
	this.RuleNames = staticData.RuleNames
	this.LiteralNames = staticData.LiteralNames
	this.SymbolicNames = staticData.SymbolicNames
	this.GrammarFileName = "VBAConditionalCompilationParser.g4"

	return this
}


// VBAConditionalCompilationParser tokens.
const (
	VBAConditionalCompilationParserEOF = antlr.TokenEOF
	VBAConditionalCompilationParserABS = 1
	VBAConditionalCompilationParserANY = 2
	VBAConditionalCompilationParserARRAY = 3
	VBAConditionalCompilationParserCBOOL = 4
	VBAConditionalCompilationParserCBYTE = 5
	VBAConditionalCompilationParserCCUR = 6
	VBAConditionalCompilationParserCDATE = 7
	VBAConditionalCompilationParserCDBL = 8
	VBAConditionalCompilationParserCDEC = 9
	VBAConditionalCompilationParserCINT = 10
	VBAConditionalCompilationParserCIRCLE = 11
	VBAConditionalCompilationParserCLNG = 12
	VBAConditionalCompilationParserCLNGLNG = 13
	VBAConditionalCompilationParserCLNGPTR = 14
	VBAConditionalCompilationParserCSNG = 15
	VBAConditionalCompilationParserCSTR = 16
	VBAConditionalCompilationParserCURRENCY = 17
	VBAConditionalCompilationParserCVAR = 18
	VBAConditionalCompilationParserCVERR = 19
	VBAConditionalCompilationParserDEBUG = 20
	VBAConditionalCompilationParserDOEVENTS = 21
	VBAConditionalCompilationParserEXIT = 22
	VBAConditionalCompilationParserFIX = 23
	VBAConditionalCompilationParserINPUTB = 24
	VBAConditionalCompilationParserINT = 25
	VBAConditionalCompilationParserLBOUND = 26
	VBAConditionalCompilationParserLEN = 27
	VBAConditionalCompilationParserLENB = 28
	VBAConditionalCompilationParserLONGLONG = 29
	VBAConditionalCompilationParserLONGPTR = 30
	VBAConditionalCompilationParserMIDB = 31
	VBAConditionalCompilationParserOPTION = 32
	VBAConditionalCompilationParserPSET = 33
	VBAConditionalCompilationParserSCALE = 34
	VBAConditionalCompilationParserSGN = 35
	VBAConditionalCompilationParserUBOUND = 36
	VBAConditionalCompilationParserCOMMA = 37
	VBAConditionalCompilationParserCOLON = 38
	VBAConditionalCompilationParserSEMICOLON = 39
	VBAConditionalCompilationParserEXCLAMATIONPOINT = 40
	VBAConditionalCompilationParserDOT = 41
	VBAConditionalCompilationParserHASH = 42
	VBAConditionalCompilationParserAT = 43
	VBAConditionalCompilationParserPERCENT = 44
	VBAConditionalCompilationParserDOLLAR = 45
	VBAConditionalCompilationParserAMPERSAND = 46
	VBAConditionalCompilationParserACCESS = 47
	VBAConditionalCompilationParserADDRESSOF = 48
	VBAConditionalCompilationParserALIAS = 49
	VBAConditionalCompilationParserAND = 50
	VBAConditionalCompilationParserATTRIBUTE = 51
	VBAConditionalCompilationParserAPPEND = 52
	VBAConditionalCompilationParserAS = 53
	VBAConditionalCompilationParserBEGINPROPERTY = 54
	VBAConditionalCompilationParserBEGIN = 55
	VBAConditionalCompilationParserBINARY = 56
	VBAConditionalCompilationParserBOOLEAN = 57
	VBAConditionalCompilationParserBYVAL = 58
	VBAConditionalCompilationParserBYREF = 59
	VBAConditionalCompilationParserBYTE = 60
	VBAConditionalCompilationParserCALL = 61
	VBAConditionalCompilationParserCASE = 62
	VBAConditionalCompilationParserCDECL = 63
	VBAConditionalCompilationParserCLASS = 64
	VBAConditionalCompilationParserCLOSE = 65
	VBAConditionalCompilationParserCONST = 66
	VBAConditionalCompilationParserDATABASE = 67
	VBAConditionalCompilationParserDATE = 68
	VBAConditionalCompilationParserDECLARE = 69
	VBAConditionalCompilationParserDEFBOOL = 70
	VBAConditionalCompilationParserDEFBYTE = 71
	VBAConditionalCompilationParserDEFDATE = 72
	VBAConditionalCompilationParserDEFDBL = 73
	VBAConditionalCompilationParserDEFCUR = 74
	VBAConditionalCompilationParserDEFINT = 75
	VBAConditionalCompilationParserDEFLNG = 76
	VBAConditionalCompilationParserDEFLNGLNG = 77
	VBAConditionalCompilationParserDEFLNGPTR = 78
	VBAConditionalCompilationParserDEFOBJ = 79
	VBAConditionalCompilationParserDEFSNG = 80
	VBAConditionalCompilationParserDEFSTR = 81
	VBAConditionalCompilationParserDEFVAR = 82
	VBAConditionalCompilationParserDIM = 83
	VBAConditionalCompilationParserDO = 84
	VBAConditionalCompilationParserDOUBLE = 85
	VBAConditionalCompilationParserEACH = 86
	VBAConditionalCompilationParserELSE = 87
	VBAConditionalCompilationParserELSEIF = 88
	VBAConditionalCompilationParserEMPTY = 89
	VBAConditionalCompilationParserEND_ENUM = 90
	VBAConditionalCompilationParserEND_FUNCTION = 91
	VBAConditionalCompilationParserEND_IF = 92
	VBAConditionalCompilationParserENDPROPERTY = 93
	VBAConditionalCompilationParserEND_PROPERTY = 94
	VBAConditionalCompilationParserEND_SELECT = 95
	VBAConditionalCompilationParserEND_SUB = 96
	VBAConditionalCompilationParserEND_TYPE = 97
	VBAConditionalCompilationParserEND_WITH = 98
	VBAConditionalCompilationParserEND = 99
	VBAConditionalCompilationParserENUM = 100
	VBAConditionalCompilationParserEQV = 101
	VBAConditionalCompilationParserERASE = 102
	VBAConditionalCompilationParserERROR = 103
	VBAConditionalCompilationParserEVENT = 104
	VBAConditionalCompilationParserEXIT_DO = 105
	VBAConditionalCompilationParserEXIT_FOR = 106
	VBAConditionalCompilationParserEXIT_FUNCTION = 107
	VBAConditionalCompilationParserEXIT_PROPERTY = 108
	VBAConditionalCompilationParserEXIT_SUB = 109
	VBAConditionalCompilationParserFALSE = 110
	VBAConditionalCompilationParserFRIEND = 111
	VBAConditionalCompilationParserFOR = 112
	VBAConditionalCompilationParserFUNCTION = 113
	VBAConditionalCompilationParserGET = 114
	VBAConditionalCompilationParserGLOBAL = 115
	VBAConditionalCompilationParserGOSUB = 116
	VBAConditionalCompilationParserGOTO = 117
	VBAConditionalCompilationParserIF = 118
	VBAConditionalCompilationParserIMP = 119
	VBAConditionalCompilationParserIMPLEMENTS = 120
	VBAConditionalCompilationParserIN = 121
	VBAConditionalCompilationParserINPUT = 122
	VBAConditionalCompilationParserIS = 123
	VBAConditionalCompilationParserINTEGER = 124
	VBAConditionalCompilationParserLOCK = 125
	VBAConditionalCompilationParserLONG = 126
	VBAConditionalCompilationParserLOOP = 127
	VBAConditionalCompilationParserLET = 128
	VBAConditionalCompilationParserLIB = 129
	VBAConditionalCompilationParserLIKE = 130
	VBAConditionalCompilationParserLINE_INPUT = 131
	VBAConditionalCompilationParserLOCK_READ = 132
	VBAConditionalCompilationParserLOCK_WRITE = 133
	VBAConditionalCompilationParserLOCK_READ_WRITE = 134
	VBAConditionalCompilationParserLSET = 135
	VBAConditionalCompilationParserME = 136
	VBAConditionalCompilationParserMID = 137
	VBAConditionalCompilationParserMOD = 138
	VBAConditionalCompilationParserNAME = 139
	VBAConditionalCompilationParserNEXT = 140
	VBAConditionalCompilationParserNEW = 141
	VBAConditionalCompilationParserNOT = 142
	VBAConditionalCompilationParserNOTHING = 143
	VBAConditionalCompilationParserNULL = 144
	VBAConditionalCompilationParserOBJECT = 145
	VBAConditionalCompilationParserON = 146
	VBAConditionalCompilationParserON_ERROR = 147
	VBAConditionalCompilationParserON_LOCAL_ERROR = 148
	VBAConditionalCompilationParserOPEN = 149
	VBAConditionalCompilationParserOPTIONAL = 150
	VBAConditionalCompilationParserOPTION_BASE = 151
	VBAConditionalCompilationParserOPTION_EXPLICIT = 152
	VBAConditionalCompilationParserOPTION_COMPARE = 153
	VBAConditionalCompilationParserOPTION_PRIVATE_MODULE = 154
	VBAConditionalCompilationParserOR = 155
	VBAConditionalCompilationParserOUTPUT = 156
	VBAConditionalCompilationParserPARAMARRAY = 157
	VBAConditionalCompilationParserPRESERVE = 158
	VBAConditionalCompilationParserPRINT = 159
	VBAConditionalCompilationParserPRIVATE = 160
	VBAConditionalCompilationParserPROPERTY_GET = 161
	VBAConditionalCompilationParserPROPERTY_LET = 162
	VBAConditionalCompilationParserPROPERTY_SET = 163
	VBAConditionalCompilationParserPTRSAFE = 164
	VBAConditionalCompilationParserPUBLIC = 165
	VBAConditionalCompilationParserPUT = 166
	VBAConditionalCompilationParserRANDOM = 167
	VBAConditionalCompilationParserRANDOMIZE = 168
	VBAConditionalCompilationParserRAISEEVENT = 169
	VBAConditionalCompilationParserREAD = 170
	VBAConditionalCompilationParserREAD_WRITE = 171
	VBAConditionalCompilationParserREDIM = 172
	VBAConditionalCompilationParserREM = 173
	VBAConditionalCompilationParserRESET = 174
	VBAConditionalCompilationParserRESUME = 175
	VBAConditionalCompilationParserRETURN = 176
	VBAConditionalCompilationParserRSET = 177
	VBAConditionalCompilationParserSEEK = 178
	VBAConditionalCompilationParserSELECT = 179
	VBAConditionalCompilationParserSET = 180
	VBAConditionalCompilationParserSHARED = 181
	VBAConditionalCompilationParserSINGLE = 182
	VBAConditionalCompilationParserSPC = 183
	VBAConditionalCompilationParserSTATIC = 184
	VBAConditionalCompilationParserSTEP = 185
	VBAConditionalCompilationParserSTOP = 186
	VBAConditionalCompilationParserSTRING = 187
	VBAConditionalCompilationParserSUB = 188
	VBAConditionalCompilationParserTAB = 189
	VBAConditionalCompilationParserTEXT = 190
	VBAConditionalCompilationParserTHEN = 191
	VBAConditionalCompilationParserTO = 192
	VBAConditionalCompilationParserTRUE = 193
	VBAConditionalCompilationParserTYPE = 194
	VBAConditionalCompilationParserTYPEOF = 195
	VBAConditionalCompilationParserUNLOCK = 196
	VBAConditionalCompilationParserUNTIL = 197
	VBAConditionalCompilationParserVARIANT = 198
	VBAConditionalCompilationParserVERSION = 199
	VBAConditionalCompilationParserWEND = 200
	VBAConditionalCompilationParserWHILE = 201
	VBAConditionalCompilationParserWIDTH = 202
	VBAConditionalCompilationParserWITH = 203
	VBAConditionalCompilationParserWITHEVENTS = 204
	VBAConditionalCompilationParserWRITE = 205
	VBAConditionalCompilationParserXOR = 206
	VBAConditionalCompilationParserASSIGN = 207
	VBAConditionalCompilationParserDIV = 208
	VBAConditionalCompilationParserINTDIV = 209
	VBAConditionalCompilationParserEQ = 210
	VBAConditionalCompilationParserGEQ = 211
	VBAConditionalCompilationParserGT = 212
	VBAConditionalCompilationParserLEQ = 213
	VBAConditionalCompilationParserLPAREN = 214
	VBAConditionalCompilationParserLT = 215
	VBAConditionalCompilationParserMINUS = 216
	VBAConditionalCompilationParserMULT = 217
	VBAConditionalCompilationParserNEQ = 218
	VBAConditionalCompilationParserPLUS = 219
	VBAConditionalCompilationParserPOW = 220
	VBAConditionalCompilationParserRPAREN = 221
	VBAConditionalCompilationParserL_SQUARE_BRACKET = 222
	VBAConditionalCompilationParserR_SQUARE_BRACKET = 223
	VBAConditionalCompilationParserL_BRACE = 224
	VBAConditionalCompilationParserR_BRACE = 225
	VBAConditionalCompilationParserSTRINGLITERAL = 226
	VBAConditionalCompilationParserOCTLITERAL = 227
	VBAConditionalCompilationParserHEXLITERAL = 228
	VBAConditionalCompilationParserFLOATLITERAL = 229
	VBAConditionalCompilationParserINTEGERLITERAL = 230
	VBAConditionalCompilationParserDATELITERAL = 231
	VBAConditionalCompilationParserNEWLINE = 232
	VBAConditionalCompilationParserSINGLEQUOTE = 233
	VBAConditionalCompilationParserUNDERSCORE = 234
	VBAConditionalCompilationParserWS = 235
	VBAConditionalCompilationParserGUIDLITERAL = 236
	VBAConditionalCompilationParserIDENTIFIER = 237
	VBAConditionalCompilationParserLINE_CONTINUATION = 238
	VBAConditionalCompilationParserBARE_HEX_LITERAL = 239
	VBAConditionalCompilationParserERRORCHAR = 240
	VBAConditionalCompilationParserLOAD = 241
	VBAConditionalCompilationParserMIDBTYPESUFFIX = 242
	VBAConditionalCompilationParserMIDTYPESUFFIX = 243
	VBAConditionalCompilationParserRESUME_NEXT = 244
)

// VBAConditionalCompilationParser rules.
const (
	VBAConditionalCompilationParserRULE_compilationUnit = 0
	VBAConditionalCompilationParserRULE_ccBlock = 1
	VBAConditionalCompilationParserRULE_ccConst = 2
	VBAConditionalCompilationParserRULE_ccVarLhs = 3
	VBAConditionalCompilationParserRULE_ccIfBlock = 4
	VBAConditionalCompilationParserRULE_ccIf = 5
	VBAConditionalCompilationParserRULE_ccElseIfBlock = 6
	VBAConditionalCompilationParserRULE_ccElseIf = 7
	VBAConditionalCompilationParserRULE_ccElseBlock = 8
	VBAConditionalCompilationParserRULE_ccElse = 9
	VBAConditionalCompilationParserRULE_ccEndIf = 10
	VBAConditionalCompilationParserRULE_ccEol = 11
	VBAConditionalCompilationParserRULE_hashConst = 12
	VBAConditionalCompilationParserRULE_hashIf = 13
	VBAConditionalCompilationParserRULE_hashElseIf = 14
	VBAConditionalCompilationParserRULE_hashElse = 15
	VBAConditionalCompilationParserRULE_hashEndIf = 16
	VBAConditionalCompilationParserRULE_physicalLine = 17
	VBAConditionalCompilationParserRULE_ccExpression = 18
	VBAConditionalCompilationParserRULE_intrinsicFunction = 19
	VBAConditionalCompilationParserRULE_intrinsicFunctionName = 20
	VBAConditionalCompilationParserRULE_name = 21
	VBAConditionalCompilationParserRULE_nameValue = 22
	VBAConditionalCompilationParserRULE_foreignName = 23
	VBAConditionalCompilationParserRULE_foreignIdentifier = 24
	VBAConditionalCompilationParserRULE_typeHint = 25
	VBAConditionalCompilationParserRULE_literal = 26
	VBAConditionalCompilationParserRULE_comment = 27
	VBAConditionalCompilationParserRULE_keyword = 28
	VBAConditionalCompilationParserRULE_markerKeyword = 29
	VBAConditionalCompilationParserRULE_statementKeyword = 30
	VBAConditionalCompilationParserRULE_whiteSpace = 31
)

// ICompilationUnitContext is an interface to support dynamic dispatch.
type ICompilationUnitContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	CcBlock() ICcBlockContext
	EOF() antlr.TerminalNode

	// IsCompilationUnitContext differentiates from other interfaces.
	IsCompilationUnitContext()
}

type CompilationUnitContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCompilationUnitContext() *CompilationUnitContext {
	var p = new(CompilationUnitContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_compilationUnit
	return p
}

func InitEmptyCompilationUnitContext(p *CompilationUnitContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_compilationUnit
}

func (*CompilationUnitContext) IsCompilationUnitContext() {}

func NewCompilationUnitContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CompilationUnitContext {
	var p = new(CompilationUnitContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_compilationUnit

	return p
}

func (s *CompilationUnitContext) GetParser() antlr.Parser { return s.parser }

func (s *CompilationUnitContext) CcBlock() ICcBlockContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcBlockContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcBlockContext)
}

func (s *CompilationUnitContext) EOF() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEOF, 0)
}

func (s *CompilationUnitContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CompilationUnitContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CompilationUnit() (localctx ICompilationUnitContext) {
	localctx = NewCompilationUnitContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 0, VBAConditionalCompilationParserRULE_compilationUnit)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(64)
		p.CcBlock()
	}
	{
		p.SetState(65)
		p.Match(VBAConditionalCompilationParserEOF)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcBlockContext is an interface to support dynamic dispatch.
type ICcBlockContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	AllCcConst() []ICcConstContext
	CcConst(i int) ICcConstContext
	AllCcIfBlock() []ICcIfBlockContext
	CcIfBlock(i int) ICcIfBlockContext
	AllPhysicalLine() []IPhysicalLineContext
	PhysicalLine(i int) IPhysicalLineContext

	// IsCcBlockContext differentiates from other interfaces.
	IsCcBlockContext()
}

type CcBlockContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcBlockContext() *CcBlockContext {
	var p = new(CcBlockContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccBlock
	return p
}

func InitEmptyCcBlockContext(p *CcBlockContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccBlock
}

func (*CcBlockContext) IsCcBlockContext() {}

func NewCcBlockContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcBlockContext {
	var p = new(CcBlockContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccBlock

	return p
}

func (s *CcBlockContext) GetParser() antlr.Parser { return s.parser }

func (s *CcBlockContext) AllCcConst() []ICcConstContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(ICcConstContext); ok {
			len++
		}
	}

	tst := make([]ICcConstContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(ICcConstContext); ok {
			tst[i] = t.(ICcConstContext)
			i++
		}
	}

	return tst
}

func (s *CcBlockContext) CcConst(i int) ICcConstContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcConstContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcConstContext)
}

func (s *CcBlockContext) AllCcIfBlock() []ICcIfBlockContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(ICcIfBlockContext); ok {
			len++
		}
	}

	tst := make([]ICcIfBlockContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(ICcIfBlockContext); ok {
			tst[i] = t.(ICcIfBlockContext)
			i++
		}
	}

	return tst
}

func (s *CcBlockContext) CcIfBlock(i int) ICcIfBlockContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcIfBlockContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcIfBlockContext)
}

func (s *CcBlockContext) AllPhysicalLine() []IPhysicalLineContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(IPhysicalLineContext); ok {
			len++
		}
	}

	tst := make([]IPhysicalLineContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(IPhysicalLineContext); ok {
			tst[i] = t.(IPhysicalLineContext)
			i++
		}
	}

	return tst
}

func (s *CcBlockContext) PhysicalLine(i int) IPhysicalLineContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IPhysicalLineContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(IPhysicalLineContext)
}

func (s *CcBlockContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcBlockContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcBlock() (localctx ICcBlockContext) {
	localctx = NewCcBlockContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 2, VBAConditionalCompilationParserRULE_ccBlock)
	var _alt int

	p.EnterOuterAlt(localctx, 1)
	p.SetState(72)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_alt = p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 1, p.GetParserRuleContext())
	if p.HasError() {
		goto errorExit
	}
	for _alt != 1 && _alt != antlr.ATNInvalidAltNumber {
		if _alt == 1+1 {
			p.SetState(70)
			p.GetErrorHandler().Sync(p)
			if p.HasError() {
				goto errorExit
			}

			switch p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 0, p.GetParserRuleContext()) {
			case 1:
				{
					p.SetState(67)
					p.CcConst()
				}


			case 2:
				{
					p.SetState(68)
					p.CcIfBlock()
				}


			case 3:
				{
					p.SetState(69)
					p.PhysicalLine()
				}

			case antlr.ATNInvalidAltNumber:
				goto errorExit
			}

		}
		p.SetState(74)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_alt = p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 1, p.GetParserRuleContext())
		if p.HasError() {
			goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcConstContext is an interface to support dynamic dispatch.
type ICcConstContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	HashConst() IHashConstContext
	CcVarLhs() ICcVarLhsContext
	EQ() antlr.TerminalNode
	CcExpression() ICcExpressionContext
	CcEol() ICcEolContext
	AllWhiteSpace() []IWhiteSpaceContext
	WhiteSpace(i int) IWhiteSpaceContext

	// IsCcConstContext differentiates from other interfaces.
	IsCcConstContext()
}

type CcConstContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcConstContext() *CcConstContext {
	var p = new(CcConstContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccConst
	return p
}

func InitEmptyCcConstContext(p *CcConstContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccConst
}

func (*CcConstContext) IsCcConstContext() {}

func NewCcConstContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcConstContext {
	var p = new(CcConstContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccConst

	return p
}

func (s *CcConstContext) GetParser() antlr.Parser { return s.parser }

func (s *CcConstContext) HashConst() IHashConstContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IHashConstContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IHashConstContext)
}

func (s *CcConstContext) CcVarLhs() ICcVarLhsContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcVarLhsContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcVarLhsContext)
}

func (s *CcConstContext) EQ() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEQ, 0)
}

func (s *CcConstContext) CcExpression() ICcExpressionContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcExpressionContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcExpressionContext)
}

func (s *CcConstContext) CcEol() ICcEolContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcEolContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcEolContext)
}

func (s *CcConstContext) AllWhiteSpace() []IWhiteSpaceContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			len++
		}
	}

	tst := make([]IWhiteSpaceContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(IWhiteSpaceContext); ok {
			tst[i] = t.(IWhiteSpaceContext)
			i++
		}
	}

	return tst
}

func (s *CcConstContext) WhiteSpace(i int) IWhiteSpaceContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(IWhiteSpaceContext)
}

func (s *CcConstContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcConstContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcConst() (localctx ICcConstContext) {
	localctx = NewCcConstContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 4, VBAConditionalCompilationParserRULE_ccConst)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	p.SetState(78)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(75)
			p.WhiteSpace()
		}


		p.SetState(80)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(81)
		p.HashConst()
	}
	p.SetState(83)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for ok := true; ok; ok = _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(82)
			p.WhiteSpace()
		}


		p.SetState(85)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(87)
		p.CcVarLhs()
	}
	p.SetState(91)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(88)
			p.WhiteSpace()
		}


		p.SetState(93)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(94)
		p.Match(VBAConditionalCompilationParserEQ)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	p.SetState(98)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(95)
			p.WhiteSpace()
		}


		p.SetState(100)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(101)
		p.ccExpression(0)
	}
	{
		p.SetState(102)
		p.CcEol()
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcVarLhsContext is an interface to support dynamic dispatch.
type ICcVarLhsContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	Name() INameContext

	// IsCcVarLhsContext differentiates from other interfaces.
	IsCcVarLhsContext()
}

type CcVarLhsContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcVarLhsContext() *CcVarLhsContext {
	var p = new(CcVarLhsContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccVarLhs
	return p
}

func InitEmptyCcVarLhsContext(p *CcVarLhsContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccVarLhs
}

func (*CcVarLhsContext) IsCcVarLhsContext() {}

func NewCcVarLhsContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcVarLhsContext {
	var p = new(CcVarLhsContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccVarLhs

	return p
}

func (s *CcVarLhsContext) GetParser() antlr.Parser { return s.parser }

func (s *CcVarLhsContext) Name() INameContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(INameContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(INameContext)
}

func (s *CcVarLhsContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcVarLhsContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcVarLhs() (localctx ICcVarLhsContext) {
	localctx = NewCcVarLhsContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 6, VBAConditionalCompilationParserRULE_ccVarLhs)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(104)
		p.Name()
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcIfBlockContext is an interface to support dynamic dispatch.
type ICcIfBlockContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	CcIf() ICcIfContext
	CcBlock() ICcBlockContext
	CcEndIf() ICcEndIfContext
	AllCcElseIfBlock() []ICcElseIfBlockContext
	CcElseIfBlock(i int) ICcElseIfBlockContext
	CcElseBlock() ICcElseBlockContext

	// IsCcIfBlockContext differentiates from other interfaces.
	IsCcIfBlockContext()
}

type CcIfBlockContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcIfBlockContext() *CcIfBlockContext {
	var p = new(CcIfBlockContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccIfBlock
	return p
}

func InitEmptyCcIfBlockContext(p *CcIfBlockContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccIfBlock
}

func (*CcIfBlockContext) IsCcIfBlockContext() {}

func NewCcIfBlockContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcIfBlockContext {
	var p = new(CcIfBlockContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccIfBlock

	return p
}

func (s *CcIfBlockContext) GetParser() antlr.Parser { return s.parser }

func (s *CcIfBlockContext) CcIf() ICcIfContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcIfContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcIfContext)
}

func (s *CcIfBlockContext) CcBlock() ICcBlockContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcBlockContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcBlockContext)
}

func (s *CcIfBlockContext) CcEndIf() ICcEndIfContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcEndIfContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcEndIfContext)
}

func (s *CcIfBlockContext) AllCcElseIfBlock() []ICcElseIfBlockContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(ICcElseIfBlockContext); ok {
			len++
		}
	}

	tst := make([]ICcElseIfBlockContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(ICcElseIfBlockContext); ok {
			tst[i] = t.(ICcElseIfBlockContext)
			i++
		}
	}

	return tst
}

func (s *CcIfBlockContext) CcElseIfBlock(i int) ICcElseIfBlockContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcElseIfBlockContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcElseIfBlockContext)
}

func (s *CcIfBlockContext) CcElseBlock() ICcElseBlockContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcElseBlockContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcElseBlockContext)
}

func (s *CcIfBlockContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcIfBlockContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcIfBlock() (localctx ICcIfBlockContext) {
	localctx = NewCcIfBlockContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 8, VBAConditionalCompilationParserRULE_ccIfBlock)
	var _alt int

	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(106)
		p.CcIf()
	}
	{
		p.SetState(107)
		p.CcBlock()
	}
	p.SetState(111)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_alt = p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 6, p.GetParserRuleContext())
	if p.HasError() {
		goto errorExit
	}
	for _alt != 2 && _alt != antlr.ATNInvalidAltNumber {
		if _alt == 1 {
			{
				p.SetState(108)
				p.CcElseIfBlock()
			}


		}
		p.SetState(113)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_alt = p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 6, p.GetParserRuleContext())
		if p.HasError() {
			goto errorExit
		}
	}
	p.SetState(115)
	p.GetErrorHandler().Sync(p)


	if p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 7, p.GetParserRuleContext()) == 1 {
		{
			p.SetState(114)
			p.CcElseBlock()
		}

		} else if p.HasError() { // JIM
			goto errorExit
	}
	{
		p.SetState(117)
		p.CcEndIf()
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcIfContext is an interface to support dynamic dispatch.
type ICcIfContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	HashIf() IHashIfContext
	CcExpression() ICcExpressionContext
	THEN() antlr.TerminalNode
	CcEol() ICcEolContext
	AllWhiteSpace() []IWhiteSpaceContext
	WhiteSpace(i int) IWhiteSpaceContext

	// IsCcIfContext differentiates from other interfaces.
	IsCcIfContext()
}

type CcIfContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcIfContext() *CcIfContext {
	var p = new(CcIfContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccIf
	return p
}

func InitEmptyCcIfContext(p *CcIfContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccIf
}

func (*CcIfContext) IsCcIfContext() {}

func NewCcIfContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcIfContext {
	var p = new(CcIfContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccIf

	return p
}

func (s *CcIfContext) GetParser() antlr.Parser { return s.parser }

func (s *CcIfContext) HashIf() IHashIfContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IHashIfContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IHashIfContext)
}

func (s *CcIfContext) CcExpression() ICcExpressionContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcExpressionContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcExpressionContext)
}

func (s *CcIfContext) THEN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserTHEN, 0)
}

func (s *CcIfContext) CcEol() ICcEolContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcEolContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcEolContext)
}

func (s *CcIfContext) AllWhiteSpace() []IWhiteSpaceContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			len++
		}
	}

	tst := make([]IWhiteSpaceContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(IWhiteSpaceContext); ok {
			tst[i] = t.(IWhiteSpaceContext)
			i++
		}
	}

	return tst
}

func (s *CcIfContext) WhiteSpace(i int) IWhiteSpaceContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(IWhiteSpaceContext)
}

func (s *CcIfContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcIfContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcIf() (localctx ICcIfContext) {
	localctx = NewCcIfContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 10, VBAConditionalCompilationParserRULE_ccIf)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	p.SetState(122)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(119)
			p.WhiteSpace()
		}


		p.SetState(124)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(125)
		p.HashIf()
	}
	p.SetState(127)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for ok := true; ok; ok = _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(126)
			p.WhiteSpace()
		}


		p.SetState(129)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(131)
		p.ccExpression(0)
	}
	p.SetState(133)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for ok := true; ok; ok = _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(132)
			p.WhiteSpace()
		}


		p.SetState(135)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(137)
		p.Match(VBAConditionalCompilationParserTHEN)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	{
		p.SetState(138)
		p.CcEol()
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcElseIfBlockContext is an interface to support dynamic dispatch.
type ICcElseIfBlockContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	CcElseIf() ICcElseIfContext
	CcBlock() ICcBlockContext

	// IsCcElseIfBlockContext differentiates from other interfaces.
	IsCcElseIfBlockContext()
}

type CcElseIfBlockContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcElseIfBlockContext() *CcElseIfBlockContext {
	var p = new(CcElseIfBlockContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElseIfBlock
	return p
}

func InitEmptyCcElseIfBlockContext(p *CcElseIfBlockContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElseIfBlock
}

func (*CcElseIfBlockContext) IsCcElseIfBlockContext() {}

func NewCcElseIfBlockContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcElseIfBlockContext {
	var p = new(CcElseIfBlockContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElseIfBlock

	return p
}

func (s *CcElseIfBlockContext) GetParser() antlr.Parser { return s.parser }

func (s *CcElseIfBlockContext) CcElseIf() ICcElseIfContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcElseIfContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcElseIfContext)
}

func (s *CcElseIfBlockContext) CcBlock() ICcBlockContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcBlockContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcBlockContext)
}

func (s *CcElseIfBlockContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcElseIfBlockContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcElseIfBlock() (localctx ICcElseIfBlockContext) {
	localctx = NewCcElseIfBlockContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 12, VBAConditionalCompilationParserRULE_ccElseIfBlock)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(140)
		p.CcElseIf()
	}
	{
		p.SetState(141)
		p.CcBlock()
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcElseIfContext is an interface to support dynamic dispatch.
type ICcElseIfContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	HashElseIf() IHashElseIfContext
	CcExpression() ICcExpressionContext
	THEN() antlr.TerminalNode
	CcEol() ICcEolContext
	AllWhiteSpace() []IWhiteSpaceContext
	WhiteSpace(i int) IWhiteSpaceContext

	// IsCcElseIfContext differentiates from other interfaces.
	IsCcElseIfContext()
}

type CcElseIfContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcElseIfContext() *CcElseIfContext {
	var p = new(CcElseIfContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElseIf
	return p
}

func InitEmptyCcElseIfContext(p *CcElseIfContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElseIf
}

func (*CcElseIfContext) IsCcElseIfContext() {}

func NewCcElseIfContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcElseIfContext {
	var p = new(CcElseIfContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElseIf

	return p
}

func (s *CcElseIfContext) GetParser() antlr.Parser { return s.parser }

func (s *CcElseIfContext) HashElseIf() IHashElseIfContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IHashElseIfContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IHashElseIfContext)
}

func (s *CcElseIfContext) CcExpression() ICcExpressionContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcExpressionContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcExpressionContext)
}

func (s *CcElseIfContext) THEN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserTHEN, 0)
}

func (s *CcElseIfContext) CcEol() ICcEolContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcEolContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcEolContext)
}

func (s *CcElseIfContext) AllWhiteSpace() []IWhiteSpaceContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			len++
		}
	}

	tst := make([]IWhiteSpaceContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(IWhiteSpaceContext); ok {
			tst[i] = t.(IWhiteSpaceContext)
			i++
		}
	}

	return tst
}

func (s *CcElseIfContext) WhiteSpace(i int) IWhiteSpaceContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(IWhiteSpaceContext)
}

func (s *CcElseIfContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcElseIfContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcElseIf() (localctx ICcElseIfContext) {
	localctx = NewCcElseIfContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 14, VBAConditionalCompilationParserRULE_ccElseIf)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	p.SetState(146)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(143)
			p.WhiteSpace()
		}


		p.SetState(148)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(149)
		p.HashElseIf()
	}
	p.SetState(151)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for ok := true; ok; ok = _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(150)
			p.WhiteSpace()
		}


		p.SetState(153)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(155)
		p.ccExpression(0)
	}
	p.SetState(157)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for ok := true; ok; ok = _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(156)
			p.WhiteSpace()
		}


		p.SetState(159)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(161)
		p.Match(VBAConditionalCompilationParserTHEN)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	{
		p.SetState(162)
		p.CcEol()
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcElseBlockContext is an interface to support dynamic dispatch.
type ICcElseBlockContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	CcElse() ICcElseContext
	CcBlock() ICcBlockContext

	// IsCcElseBlockContext differentiates from other interfaces.
	IsCcElseBlockContext()
}

type CcElseBlockContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcElseBlockContext() *CcElseBlockContext {
	var p = new(CcElseBlockContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElseBlock
	return p
}

func InitEmptyCcElseBlockContext(p *CcElseBlockContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElseBlock
}

func (*CcElseBlockContext) IsCcElseBlockContext() {}

func NewCcElseBlockContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcElseBlockContext {
	var p = new(CcElseBlockContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElseBlock

	return p
}

func (s *CcElseBlockContext) GetParser() antlr.Parser { return s.parser }

func (s *CcElseBlockContext) CcElse() ICcElseContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcElseContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcElseContext)
}

func (s *CcElseBlockContext) CcBlock() ICcBlockContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcBlockContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcBlockContext)
}

func (s *CcElseBlockContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcElseBlockContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcElseBlock() (localctx ICcElseBlockContext) {
	localctx = NewCcElseBlockContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 16, VBAConditionalCompilationParserRULE_ccElseBlock)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(164)
		p.CcElse()
	}
	{
		p.SetState(165)
		p.CcBlock()
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcElseContext is an interface to support dynamic dispatch.
type ICcElseContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	HashElse() IHashElseContext
	CcEol() ICcEolContext
	AllWhiteSpace() []IWhiteSpaceContext
	WhiteSpace(i int) IWhiteSpaceContext

	// IsCcElseContext differentiates from other interfaces.
	IsCcElseContext()
}

type CcElseContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcElseContext() *CcElseContext {
	var p = new(CcElseContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElse
	return p
}

func InitEmptyCcElseContext(p *CcElseContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElse
}

func (*CcElseContext) IsCcElseContext() {}

func NewCcElseContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcElseContext {
	var p = new(CcElseContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccElse

	return p
}

func (s *CcElseContext) GetParser() antlr.Parser { return s.parser }

func (s *CcElseContext) HashElse() IHashElseContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IHashElseContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IHashElseContext)
}

func (s *CcElseContext) CcEol() ICcEolContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcEolContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcEolContext)
}

func (s *CcElseContext) AllWhiteSpace() []IWhiteSpaceContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			len++
		}
	}

	tst := make([]IWhiteSpaceContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(IWhiteSpaceContext); ok {
			tst[i] = t.(IWhiteSpaceContext)
			i++
		}
	}

	return tst
}

func (s *CcElseContext) WhiteSpace(i int) IWhiteSpaceContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(IWhiteSpaceContext)
}

func (s *CcElseContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcElseContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcElse() (localctx ICcElseContext) {
	localctx = NewCcElseContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 18, VBAConditionalCompilationParserRULE_ccElse)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	p.SetState(170)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(167)
			p.WhiteSpace()
		}


		p.SetState(172)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(173)
		p.HashElse()
	}
	{
		p.SetState(174)
		p.CcEol()
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcEndIfContext is an interface to support dynamic dispatch.
type ICcEndIfContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	HashEndIf() IHashEndIfContext
	CcEol() ICcEolContext
	AllWhiteSpace() []IWhiteSpaceContext
	WhiteSpace(i int) IWhiteSpaceContext

	// IsCcEndIfContext differentiates from other interfaces.
	IsCcEndIfContext()
}

type CcEndIfContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcEndIfContext() *CcEndIfContext {
	var p = new(CcEndIfContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccEndIf
	return p
}

func InitEmptyCcEndIfContext(p *CcEndIfContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccEndIf
}

func (*CcEndIfContext) IsCcEndIfContext() {}

func NewCcEndIfContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcEndIfContext {
	var p = new(CcEndIfContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccEndIf

	return p
}

func (s *CcEndIfContext) GetParser() antlr.Parser { return s.parser }

func (s *CcEndIfContext) HashEndIf() IHashEndIfContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IHashEndIfContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IHashEndIfContext)
}

func (s *CcEndIfContext) CcEol() ICcEolContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcEolContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcEolContext)
}

func (s *CcEndIfContext) AllWhiteSpace() []IWhiteSpaceContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			len++
		}
	}

	tst := make([]IWhiteSpaceContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(IWhiteSpaceContext); ok {
			tst[i] = t.(IWhiteSpaceContext)
			i++
		}
	}

	return tst
}

func (s *CcEndIfContext) WhiteSpace(i int) IWhiteSpaceContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(IWhiteSpaceContext)
}

func (s *CcEndIfContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcEndIfContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcEndIf() (localctx ICcEndIfContext) {
	localctx = NewCcEndIfContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 20, VBAConditionalCompilationParserRULE_ccEndIf)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	p.SetState(179)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(176)
			p.WhiteSpace()
		}


		p.SetState(181)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(182)
		p.HashEndIf()
	}
	{
		p.SetState(183)
		p.CcEol()
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcEolContext is an interface to support dynamic dispatch.
type ICcEolContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	NEWLINE() antlr.TerminalNode
	AllWhiteSpace() []IWhiteSpaceContext
	WhiteSpace(i int) IWhiteSpaceContext
	Comment() ICommentContext

	// IsCcEolContext differentiates from other interfaces.
	IsCcEolContext()
}

type CcEolContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcEolContext() *CcEolContext {
	var p = new(CcEolContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccEol
	return p
}

func InitEmptyCcEolContext(p *CcEolContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccEol
}

func (*CcEolContext) IsCcEolContext() {}

func NewCcEolContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcEolContext {
	var p = new(CcEolContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccEol

	return p
}

func (s *CcEolContext) GetParser() antlr.Parser { return s.parser }

func (s *CcEolContext) NEWLINE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNEWLINE, 0)
}

func (s *CcEolContext) AllWhiteSpace() []IWhiteSpaceContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			len++
		}
	}

	tst := make([]IWhiteSpaceContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(IWhiteSpaceContext); ok {
			tst[i] = t.(IWhiteSpaceContext)
			i++
		}
	}

	return tst
}

func (s *CcEolContext) WhiteSpace(i int) IWhiteSpaceContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(IWhiteSpaceContext)
}

func (s *CcEolContext) Comment() ICommentContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICommentContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICommentContext)
}

func (s *CcEolContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcEolContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) CcEol() (localctx ICcEolContext) {
	localctx = NewCcEolContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 22, VBAConditionalCompilationParserRULE_ccEol)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	p.SetState(188)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(185)
			p.WhiteSpace()
		}


		p.SetState(190)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	p.SetState(192)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	if _la == VBAConditionalCompilationParserSINGLEQUOTE {
		{
			p.SetState(191)
			p.Comment()
		}

	}
	{
		p.SetState(194)
		p.Match(VBAConditionalCompilationParserNEWLINE)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IHashConstContext is an interface to support dynamic dispatch.
type IHashConstContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	HASH() antlr.TerminalNode
	CONST() antlr.TerminalNode

	// IsHashConstContext differentiates from other interfaces.
	IsHashConstContext()
}

type HashConstContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyHashConstContext() *HashConstContext {
	var p = new(HashConstContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashConst
	return p
}

func InitEmptyHashConstContext(p *HashConstContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashConst
}

func (*HashConstContext) IsHashConstContext() {}

func NewHashConstContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *HashConstContext {
	var p = new(HashConstContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashConst

	return p
}

func (s *HashConstContext) GetParser() antlr.Parser { return s.parser }

func (s *HashConstContext) HASH() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserHASH, 0)
}

func (s *HashConstContext) CONST() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCONST, 0)
}

func (s *HashConstContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *HashConstContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) HashConst() (localctx IHashConstContext) {
	localctx = NewHashConstContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 24, VBAConditionalCompilationParserRULE_hashConst)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(196)
		p.Match(VBAConditionalCompilationParserHASH)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	{
		p.SetState(197)
		p.Match(VBAConditionalCompilationParserCONST)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IHashIfContext is an interface to support dynamic dispatch.
type IHashIfContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	HASH() antlr.TerminalNode
	IF() antlr.TerminalNode

	// IsHashIfContext differentiates from other interfaces.
	IsHashIfContext()
}

type HashIfContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyHashIfContext() *HashIfContext {
	var p = new(HashIfContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashIf
	return p
}

func InitEmptyHashIfContext(p *HashIfContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashIf
}

func (*HashIfContext) IsHashIfContext() {}

func NewHashIfContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *HashIfContext {
	var p = new(HashIfContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashIf

	return p
}

func (s *HashIfContext) GetParser() antlr.Parser { return s.parser }

func (s *HashIfContext) HASH() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserHASH, 0)
}

func (s *HashIfContext) IF() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserIF, 0)
}

func (s *HashIfContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *HashIfContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) HashIf() (localctx IHashIfContext) {
	localctx = NewHashIfContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 26, VBAConditionalCompilationParserRULE_hashIf)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(199)
		p.Match(VBAConditionalCompilationParserHASH)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	{
		p.SetState(200)
		p.Match(VBAConditionalCompilationParserIF)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IHashElseIfContext is an interface to support dynamic dispatch.
type IHashElseIfContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	HASH() antlr.TerminalNode
	ELSEIF() antlr.TerminalNode

	// IsHashElseIfContext differentiates from other interfaces.
	IsHashElseIfContext()
}

type HashElseIfContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyHashElseIfContext() *HashElseIfContext {
	var p = new(HashElseIfContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashElseIf
	return p
}

func InitEmptyHashElseIfContext(p *HashElseIfContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashElseIf
}

func (*HashElseIfContext) IsHashElseIfContext() {}

func NewHashElseIfContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *HashElseIfContext {
	var p = new(HashElseIfContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashElseIf

	return p
}

func (s *HashElseIfContext) GetParser() antlr.Parser { return s.parser }

func (s *HashElseIfContext) HASH() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserHASH, 0)
}

func (s *HashElseIfContext) ELSEIF() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserELSEIF, 0)
}

func (s *HashElseIfContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *HashElseIfContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) HashElseIf() (localctx IHashElseIfContext) {
	localctx = NewHashElseIfContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 28, VBAConditionalCompilationParserRULE_hashElseIf)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(202)
		p.Match(VBAConditionalCompilationParserHASH)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	{
		p.SetState(203)
		p.Match(VBAConditionalCompilationParserELSEIF)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IHashElseContext is an interface to support dynamic dispatch.
type IHashElseContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	HASH() antlr.TerminalNode
	ELSE() antlr.TerminalNode

	// IsHashElseContext differentiates from other interfaces.
	IsHashElseContext()
}

type HashElseContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyHashElseContext() *HashElseContext {
	var p = new(HashElseContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashElse
	return p
}

func InitEmptyHashElseContext(p *HashElseContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashElse
}

func (*HashElseContext) IsHashElseContext() {}

func NewHashElseContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *HashElseContext {
	var p = new(HashElseContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashElse

	return p
}

func (s *HashElseContext) GetParser() antlr.Parser { return s.parser }

func (s *HashElseContext) HASH() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserHASH, 0)
}

func (s *HashElseContext) ELSE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserELSE, 0)
}

func (s *HashElseContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *HashElseContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) HashElse() (localctx IHashElseContext) {
	localctx = NewHashElseContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 30, VBAConditionalCompilationParserRULE_hashElse)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(205)
		p.Match(VBAConditionalCompilationParserHASH)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	{
		p.SetState(206)
		p.Match(VBAConditionalCompilationParserELSE)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IHashEndIfContext is an interface to support dynamic dispatch.
type IHashEndIfContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	HASH() antlr.TerminalNode
	END_IF() antlr.TerminalNode

	// IsHashEndIfContext differentiates from other interfaces.
	IsHashEndIfContext()
}

type HashEndIfContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyHashEndIfContext() *HashEndIfContext {
	var p = new(HashEndIfContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashEndIf
	return p
}

func InitEmptyHashEndIfContext(p *HashEndIfContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashEndIf
}

func (*HashEndIfContext) IsHashEndIfContext() {}

func NewHashEndIfContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *HashEndIfContext {
	var p = new(HashEndIfContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_hashEndIf

	return p
}

func (s *HashEndIfContext) GetParser() antlr.Parser { return s.parser }

func (s *HashEndIfContext) HASH() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserHASH, 0)
}

func (s *HashEndIfContext) END_IF() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEND_IF, 0)
}

func (s *HashEndIfContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *HashEndIfContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) HashEndIf() (localctx IHashEndIfContext) {
	localctx = NewHashEndIfContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 32, VBAConditionalCompilationParserRULE_hashEndIf)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(208)
		p.Match(VBAConditionalCompilationParserHASH)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	{
		p.SetState(209)
		p.Match(VBAConditionalCompilationParserEND_IF)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IPhysicalLineContext is an interface to support dynamic dispatch.
type IPhysicalLineContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	AllNEWLINE() []antlr.TerminalNode
	NEWLINE(i int) antlr.TerminalNode

	// IsPhysicalLineContext differentiates from other interfaces.
	IsPhysicalLineContext()
}

type PhysicalLineContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyPhysicalLineContext() *PhysicalLineContext {
	var p = new(PhysicalLineContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_physicalLine
	return p
}

func InitEmptyPhysicalLineContext(p *PhysicalLineContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_physicalLine
}

func (*PhysicalLineContext) IsPhysicalLineContext() {}

func NewPhysicalLineContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *PhysicalLineContext {
	var p = new(PhysicalLineContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_physicalLine

	return p
}

func (s *PhysicalLineContext) GetParser() antlr.Parser { return s.parser }

func (s *PhysicalLineContext) AllNEWLINE() []antlr.TerminalNode {
	return s.GetTokens(VBAConditionalCompilationParserNEWLINE)
}

func (s *PhysicalLineContext) NEWLINE(i int) antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNEWLINE, i)
}

func (s *PhysicalLineContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *PhysicalLineContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) PhysicalLine() (localctx IPhysicalLineContext) {
	localctx = NewPhysicalLineContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 34, VBAConditionalCompilationParserRULE_physicalLine)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	p.SetState(214)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for ((int64(_la) & ^0x3f) == 0 && ((int64(1) << _la) & -2) != 0) || ((int64((_la - 64)) & ^0x3f) == 0 && ((int64(1) << (_la - 64)) & -1) != 0) || ((int64((_la - 128)) & ^0x3f) == 0 && ((int64(1) << (_la - 128)) & -1) != 0) || ((int64((_la - 192)) & ^0x3f) == 0 && ((int64(1) << (_la - 192)) & 9006099743113215) != 0) {
		{
			p.SetState(211)
			_la = p.GetTokenStream().LA(1)

			if _la <= 0 || _la == VBAConditionalCompilationParserNEWLINE  {
				p.GetErrorHandler().RecoverInline(p)
			} else {
				p.GetErrorHandler().ReportMatch(p)
				p.Consume()
			}
		}


		p.SetState(216)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(217)
		p.Match(VBAConditionalCompilationParserNEWLINE)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICcExpressionContext is an interface to support dynamic dispatch.
type ICcExpressionContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	LPAREN() antlr.TerminalNode
	AllCcExpression() []ICcExpressionContext
	CcExpression(i int) ICcExpressionContext
	RPAREN() antlr.TerminalNode
	AllWhiteSpace() []IWhiteSpaceContext
	WhiteSpace(i int) IWhiteSpaceContext
	MINUS() antlr.TerminalNode
	NOT() antlr.TerminalNode
	IntrinsicFunction() IIntrinsicFunctionContext
	Literal() ILiteralContext
	Name() INameContext
	POW() antlr.TerminalNode
	MULT() antlr.TerminalNode
	DIV() antlr.TerminalNode
	INTDIV() antlr.TerminalNode
	MOD() antlr.TerminalNode
	PLUS() antlr.TerminalNode
	AMPERSAND() antlr.TerminalNode
	EQ() antlr.TerminalNode
	NEQ() antlr.TerminalNode
	LT() antlr.TerminalNode
	GT() antlr.TerminalNode
	LEQ() antlr.TerminalNode
	GEQ() antlr.TerminalNode
	LIKE() antlr.TerminalNode
	IS() antlr.TerminalNode
	AND() antlr.TerminalNode
	OR() antlr.TerminalNode
	XOR() antlr.TerminalNode
	EQV() antlr.TerminalNode
	IMP() antlr.TerminalNode

	// IsCcExpressionContext differentiates from other interfaces.
	IsCcExpressionContext()
}

type CcExpressionContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCcExpressionContext() *CcExpressionContext {
	var p = new(CcExpressionContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccExpression
	return p
}

func InitEmptyCcExpressionContext(p *CcExpressionContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccExpression
}

func (*CcExpressionContext) IsCcExpressionContext() {}

func NewCcExpressionContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CcExpressionContext {
	var p = new(CcExpressionContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_ccExpression

	return p
}

func (s *CcExpressionContext) GetParser() antlr.Parser { return s.parser }

func (s *CcExpressionContext) LPAREN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLPAREN, 0)
}

func (s *CcExpressionContext) AllCcExpression() []ICcExpressionContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(ICcExpressionContext); ok {
			len++
		}
	}

	tst := make([]ICcExpressionContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(ICcExpressionContext); ok {
			tst[i] = t.(ICcExpressionContext)
			i++
		}
	}

	return tst
}

func (s *CcExpressionContext) CcExpression(i int) ICcExpressionContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcExpressionContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcExpressionContext)
}

func (s *CcExpressionContext) RPAREN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserRPAREN, 0)
}

func (s *CcExpressionContext) AllWhiteSpace() []IWhiteSpaceContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			len++
		}
	}

	tst := make([]IWhiteSpaceContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(IWhiteSpaceContext); ok {
			tst[i] = t.(IWhiteSpaceContext)
			i++
		}
	}

	return tst
}

func (s *CcExpressionContext) WhiteSpace(i int) IWhiteSpaceContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(IWhiteSpaceContext)
}

func (s *CcExpressionContext) MINUS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserMINUS, 0)
}

func (s *CcExpressionContext) NOT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNOT, 0)
}

func (s *CcExpressionContext) IntrinsicFunction() IIntrinsicFunctionContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IIntrinsicFunctionContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IIntrinsicFunctionContext)
}

func (s *CcExpressionContext) Literal() ILiteralContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ILiteralContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ILiteralContext)
}

func (s *CcExpressionContext) Name() INameContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(INameContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(INameContext)
}

func (s *CcExpressionContext) POW() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPOW, 0)
}

func (s *CcExpressionContext) MULT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserMULT, 0)
}

func (s *CcExpressionContext) DIV() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDIV, 0)
}

func (s *CcExpressionContext) INTDIV() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserINTDIV, 0)
}

func (s *CcExpressionContext) MOD() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserMOD, 0)
}

func (s *CcExpressionContext) PLUS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPLUS, 0)
}

func (s *CcExpressionContext) AMPERSAND() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserAMPERSAND, 0)
}

func (s *CcExpressionContext) EQ() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEQ, 0)
}

func (s *CcExpressionContext) NEQ() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNEQ, 0)
}

func (s *CcExpressionContext) LT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLT, 0)
}

func (s *CcExpressionContext) GT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserGT, 0)
}

func (s *CcExpressionContext) LEQ() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLEQ, 0)
}

func (s *CcExpressionContext) GEQ() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserGEQ, 0)
}

func (s *CcExpressionContext) LIKE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLIKE, 0)
}

func (s *CcExpressionContext) IS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserIS, 0)
}

func (s *CcExpressionContext) AND() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserAND, 0)
}

func (s *CcExpressionContext) OR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserOR, 0)
}

func (s *CcExpressionContext) XOR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserXOR, 0)
}

func (s *CcExpressionContext) EQV() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEQV, 0)
}

func (s *CcExpressionContext) IMP() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserIMP, 0)
}

func (s *CcExpressionContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CcExpressionContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}





func (p *VBAConditionalCompilationParser) CcExpression() (localctx ICcExpressionContext) {
	return p.ccExpression(0)
}

func (p *VBAConditionalCompilationParser) ccExpression(_p int) (localctx ICcExpressionContext) {
	var _parentctx antlr.ParserRuleContext = p.GetParserRuleContext()

	_parentState := p.GetState()
	localctx = NewCcExpressionContext(p, p.GetParserRuleContext(), _parentState)
	var _prevctx ICcExpressionContext = localctx
	var _ antlr.ParserRuleContext = _prevctx // TODO: To prevent unused variable warning.
	_startState := 36
	p.EnterRecursionRule(localctx, 36, VBAConditionalCompilationParserRULE_ccExpression, _p)
	var _la int

	var _alt int

	p.EnterOuterAlt(localctx, 1)
	p.SetState(255)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}

	switch p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 23, p.GetParserRuleContext()) {
	case 1:
		{
			p.SetState(220)
			p.Match(VBAConditionalCompilationParserLPAREN)
			if p.HasError() {
					// Recognition error - abort rule
					goto errorExit
			}
		}
		p.SetState(224)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
			goto errorExit
		}
		_la = p.GetTokenStream().LA(1)


		for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
			{
				p.SetState(221)
				p.WhiteSpace()
			}


			p.SetState(226)
			p.GetErrorHandler().Sync(p)
			if p.HasError() {
		    	goto errorExit
		    }
			_la = p.GetTokenStream().LA(1)
		}
		{
			p.SetState(227)
			p.ccExpression(0)
		}
		p.SetState(231)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
			goto errorExit
		}
		_la = p.GetTokenStream().LA(1)


		for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
			{
				p.SetState(228)
				p.WhiteSpace()
			}


			p.SetState(233)
			p.GetErrorHandler().Sync(p)
			if p.HasError() {
		    	goto errorExit
		    }
			_la = p.GetTokenStream().LA(1)
		}
		{
			p.SetState(234)
			p.Match(VBAConditionalCompilationParserRPAREN)
			if p.HasError() {
					// Recognition error - abort rule
					goto errorExit
			}
		}


	case 2:
		{
			p.SetState(236)
			p.Match(VBAConditionalCompilationParserMINUS)
			if p.HasError() {
					// Recognition error - abort rule
					goto errorExit
			}
		}
		p.SetState(240)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
			goto errorExit
		}
		_la = p.GetTokenStream().LA(1)


		for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
			{
				p.SetState(237)
				p.WhiteSpace()
			}


			p.SetState(242)
			p.GetErrorHandler().Sync(p)
			if p.HasError() {
		    	goto errorExit
		    }
			_la = p.GetTokenStream().LA(1)
		}
		{
			p.SetState(243)
			p.ccExpression(16)
		}


	case 3:
		{
			p.SetState(244)
			p.Match(VBAConditionalCompilationParserNOT)
			if p.HasError() {
					// Recognition error - abort rule
					goto errorExit
			}
		}
		p.SetState(248)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
			goto errorExit
		}
		_la = p.GetTokenStream().LA(1)


		for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
			{
				p.SetState(245)
				p.WhiteSpace()
			}


			p.SetState(250)
			p.GetErrorHandler().Sync(p)
			if p.HasError() {
		    	goto errorExit
		    }
			_la = p.GetTokenStream().LA(1)
		}
		{
			p.SetState(251)
			p.ccExpression(9)
		}


	case 4:
		{
			p.SetState(252)
			p.IntrinsicFunction()
		}


	case 5:
		{
			p.SetState(253)
			p.Literal()
		}


	case 6:
		{
			p.SetState(254)
			p.Name()
		}

	case antlr.ATNInvalidAltNumber:
		goto errorExit
	}
	p.GetParserRuleContext().SetStop(p.GetTokenStream().LT(-1))
	p.SetState(439)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_alt = p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 49, p.GetParserRuleContext())
	if p.HasError() {
		goto errorExit
	}
	for _alt != 2 && _alt != antlr.ATNInvalidAltNumber {
		if _alt == 1 {
			if p.GetParseListeners() != nil {
				p.TriggerExitRuleEvent()
			}
			_prevctx = localctx
			p.SetState(437)
			p.GetErrorHandler().Sync(p)
			if p.HasError() {
				goto errorExit
			}

			switch p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 48, p.GetParserRuleContext()) {
			case 1:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(257)

				if !(p.Precpred(p.GetParserRuleContext(), 17)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 17)", ""))
					goto errorExit
				}
				p.SetState(261)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(258)
						p.WhiteSpace()
					}


					p.SetState(263)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(264)
					p.Match(VBAConditionalCompilationParserPOW)
					if p.HasError() {
							// Recognition error - abort rule
							goto errorExit
					}
				}
				p.SetState(268)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(265)
						p.WhiteSpace()
					}


					p.SetState(270)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(271)
					p.ccExpression(18)
				}


			case 2:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(272)

				if !(p.Precpred(p.GetParserRuleContext(), 15)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 15)", ""))
					goto errorExit
				}
				p.SetState(276)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(273)
						p.WhiteSpace()
					}


					p.SetState(278)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(279)
					_la = p.GetTokenStream().LA(1)

					if !(_la == VBAConditionalCompilationParserDIV || _la == VBAConditionalCompilationParserMULT) {
						p.GetErrorHandler().RecoverInline(p)
					} else {
						p.GetErrorHandler().ReportMatch(p)
						p.Consume()
					}
				}
				p.SetState(283)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(280)
						p.WhiteSpace()
					}


					p.SetState(285)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(286)
					p.ccExpression(16)
				}


			case 3:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(287)

				if !(p.Precpred(p.GetParserRuleContext(), 14)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 14)", ""))
					goto errorExit
				}
				p.SetState(291)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(288)
						p.WhiteSpace()
					}


					p.SetState(293)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(294)
					p.Match(VBAConditionalCompilationParserINTDIV)
					if p.HasError() {
							// Recognition error - abort rule
							goto errorExit
					}
				}
				p.SetState(298)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(295)
						p.WhiteSpace()
					}


					p.SetState(300)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(301)
					p.ccExpression(15)
				}


			case 4:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(302)

				if !(p.Precpred(p.GetParserRuleContext(), 13)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 13)", ""))
					goto errorExit
				}
				p.SetState(306)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(303)
						p.WhiteSpace()
					}


					p.SetState(308)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(309)
					p.Match(VBAConditionalCompilationParserMOD)
					if p.HasError() {
							// Recognition error - abort rule
							goto errorExit
					}
				}
				p.SetState(313)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(310)
						p.WhiteSpace()
					}


					p.SetState(315)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(316)
					p.ccExpression(14)
				}


			case 5:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(317)

				if !(p.Precpred(p.GetParserRuleContext(), 12)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 12)", ""))
					goto errorExit
				}
				p.SetState(321)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(318)
						p.WhiteSpace()
					}


					p.SetState(323)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(324)
					_la = p.GetTokenStream().LA(1)

					if !(_la == VBAConditionalCompilationParserMINUS || _la == VBAConditionalCompilationParserPLUS) {
						p.GetErrorHandler().RecoverInline(p)
					} else {
						p.GetErrorHandler().ReportMatch(p)
						p.Consume()
					}
				}
				p.SetState(328)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(325)
						p.WhiteSpace()
					}


					p.SetState(330)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(331)
					p.ccExpression(13)
				}


			case 6:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(332)

				if !(p.Precpred(p.GetParserRuleContext(), 11)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 11)", ""))
					goto errorExit
				}
				p.SetState(336)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(333)
						p.WhiteSpace()
					}


					p.SetState(338)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(339)
					p.Match(VBAConditionalCompilationParserAMPERSAND)
					if p.HasError() {
							// Recognition error - abort rule
							goto errorExit
					}
				}
				p.SetState(343)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(340)
						p.WhiteSpace()
					}


					p.SetState(345)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(346)
					p.ccExpression(12)
				}


			case 7:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(347)

				if !(p.Precpred(p.GetParserRuleContext(), 10)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 10)", ""))
					goto errorExit
				}
				p.SetState(351)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(348)
						p.WhiteSpace()
					}


					p.SetState(353)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(354)
					_la = p.GetTokenStream().LA(1)

					if !(_la == VBAConditionalCompilationParserIS || _la == VBAConditionalCompilationParserLIKE || ((int64((_la - 210)) & ^0x3f) == 0 && ((int64(1) << (_la - 210)) & 303) != 0)) {
						p.GetErrorHandler().RecoverInline(p)
					} else {
						p.GetErrorHandler().ReportMatch(p)
						p.Consume()
					}
				}
				p.SetState(358)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(355)
						p.WhiteSpace()
					}


					p.SetState(360)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(361)
					p.ccExpression(11)
				}


			case 8:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(362)

				if !(p.Precpred(p.GetParserRuleContext(), 8)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 8)", ""))
					goto errorExit
				}
				p.SetState(366)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(363)
						p.WhiteSpace()
					}


					p.SetState(368)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(369)
					p.Match(VBAConditionalCompilationParserAND)
					if p.HasError() {
							// Recognition error - abort rule
							goto errorExit
					}
				}
				p.SetState(373)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(370)
						p.WhiteSpace()
					}


					p.SetState(375)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(376)
					p.ccExpression(9)
				}


			case 9:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(377)

				if !(p.Precpred(p.GetParserRuleContext(), 7)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 7)", ""))
					goto errorExit
				}
				p.SetState(381)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(378)
						p.WhiteSpace()
					}


					p.SetState(383)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(384)
					p.Match(VBAConditionalCompilationParserOR)
					if p.HasError() {
							// Recognition error - abort rule
							goto errorExit
					}
				}
				p.SetState(388)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(385)
						p.WhiteSpace()
					}


					p.SetState(390)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(391)
					p.ccExpression(8)
				}


			case 10:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(392)

				if !(p.Precpred(p.GetParserRuleContext(), 6)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 6)", ""))
					goto errorExit
				}
				p.SetState(396)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(393)
						p.WhiteSpace()
					}


					p.SetState(398)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(399)
					p.Match(VBAConditionalCompilationParserXOR)
					if p.HasError() {
							// Recognition error - abort rule
							goto errorExit
					}
				}
				p.SetState(403)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(400)
						p.WhiteSpace()
					}


					p.SetState(405)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(406)
					p.ccExpression(7)
				}


			case 11:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(407)

				if !(p.Precpred(p.GetParserRuleContext(), 5)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 5)", ""))
					goto errorExit
				}
				p.SetState(411)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(408)
						p.WhiteSpace()
					}


					p.SetState(413)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(414)
					p.Match(VBAConditionalCompilationParserEQV)
					if p.HasError() {
							// Recognition error - abort rule
							goto errorExit
					}
				}
				p.SetState(418)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(415)
						p.WhiteSpace()
					}


					p.SetState(420)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(421)
					p.ccExpression(6)
				}


			case 12:
				localctx = NewCcExpressionContext(p, _parentctx, _parentState)
				p.PushNewRecursionContext(localctx, _startState, VBAConditionalCompilationParserRULE_ccExpression)
				p.SetState(422)

				if !(p.Precpred(p.GetParserRuleContext(), 4)) {
					p.SetError(antlr.NewFailedPredicateException(p, "p.Precpred(p.GetParserRuleContext(), 4)", ""))
					goto errorExit
				}
				p.SetState(426)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(423)
						p.WhiteSpace()
					}


					p.SetState(428)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(429)
					p.Match(VBAConditionalCompilationParserIMP)
					if p.HasError() {
							// Recognition error - abort rule
							goto errorExit
					}
				}
				p.SetState(433)
				p.GetErrorHandler().Sync(p)
				if p.HasError() {
					goto errorExit
				}
				_la = p.GetTokenStream().LA(1)


				for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
					{
						p.SetState(430)
						p.WhiteSpace()
					}


					p.SetState(435)
					p.GetErrorHandler().Sync(p)
					if p.HasError() {
				    	goto errorExit
				    }
					_la = p.GetTokenStream().LA(1)
				}
				{
					p.SetState(436)
					p.ccExpression(5)
				}

			case antlr.ATNInvalidAltNumber:
				goto errorExit
			}

		}
		p.SetState(441)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_alt = p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 49, p.GetParserRuleContext())
		if p.HasError() {
			goto errorExit
		}
	}



	errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.UnrollRecursionContexts(_parentctx)
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IIntrinsicFunctionContext is an interface to support dynamic dispatch.
type IIntrinsicFunctionContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	IntrinsicFunctionName() IIntrinsicFunctionNameContext
	LPAREN() antlr.TerminalNode
	CcExpression() ICcExpressionContext
	RPAREN() antlr.TerminalNode
	AllWhiteSpace() []IWhiteSpaceContext
	WhiteSpace(i int) IWhiteSpaceContext

	// IsIntrinsicFunctionContext differentiates from other interfaces.
	IsIntrinsicFunctionContext()
}

type IntrinsicFunctionContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyIntrinsicFunctionContext() *IntrinsicFunctionContext {
	var p = new(IntrinsicFunctionContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_intrinsicFunction
	return p
}

func InitEmptyIntrinsicFunctionContext(p *IntrinsicFunctionContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_intrinsicFunction
}

func (*IntrinsicFunctionContext) IsIntrinsicFunctionContext() {}

func NewIntrinsicFunctionContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *IntrinsicFunctionContext {
	var p = new(IntrinsicFunctionContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_intrinsicFunction

	return p
}

func (s *IntrinsicFunctionContext) GetParser() antlr.Parser { return s.parser }

func (s *IntrinsicFunctionContext) IntrinsicFunctionName() IIntrinsicFunctionNameContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IIntrinsicFunctionNameContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IIntrinsicFunctionNameContext)
}

func (s *IntrinsicFunctionContext) LPAREN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLPAREN, 0)
}

func (s *IntrinsicFunctionContext) CcExpression() ICcExpressionContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ICcExpressionContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ICcExpressionContext)
}

func (s *IntrinsicFunctionContext) RPAREN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserRPAREN, 0)
}

func (s *IntrinsicFunctionContext) AllWhiteSpace() []IWhiteSpaceContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			len++
		}
	}

	tst := make([]IWhiteSpaceContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(IWhiteSpaceContext); ok {
			tst[i] = t.(IWhiteSpaceContext)
			i++
		}
	}

	return tst
}

func (s *IntrinsicFunctionContext) WhiteSpace(i int) IWhiteSpaceContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IWhiteSpaceContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(IWhiteSpaceContext)
}

func (s *IntrinsicFunctionContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *IntrinsicFunctionContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) IntrinsicFunction() (localctx IIntrinsicFunctionContext) {
	localctx = NewIntrinsicFunctionContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 38, VBAConditionalCompilationParserRULE_intrinsicFunction)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(442)
		p.IntrinsicFunctionName()
	}
	{
		p.SetState(443)
		p.Match(VBAConditionalCompilationParserLPAREN)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	p.SetState(447)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(444)
			p.WhiteSpace()
		}


		p.SetState(449)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(450)
		p.ccExpression(0)
	}
	p.SetState(454)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for _la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION {
		{
			p.SetState(451)
			p.WhiteSpace()
		}


		p.SetState(456)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}
	{
		p.SetState(457)
		p.Match(VBAConditionalCompilationParserRPAREN)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IIntrinsicFunctionNameContext is an interface to support dynamic dispatch.
type IIntrinsicFunctionNameContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	INT() antlr.TerminalNode
	FIX() antlr.TerminalNode
	ABS() antlr.TerminalNode
	SGN() antlr.TerminalNode
	LEN() antlr.TerminalNode
	LENB() antlr.TerminalNode
	CBOOL() antlr.TerminalNode
	CBYTE() antlr.TerminalNode
	CCUR() antlr.TerminalNode
	CDATE() antlr.TerminalNode
	CDBL() antlr.TerminalNode
	CINT() antlr.TerminalNode
	CLNG() antlr.TerminalNode
	CLNGLNG() antlr.TerminalNode
	CLNGPTR() antlr.TerminalNode
	CSNG() antlr.TerminalNode
	CSTR() antlr.TerminalNode
	CVAR() antlr.TerminalNode

	// IsIntrinsicFunctionNameContext differentiates from other interfaces.
	IsIntrinsicFunctionNameContext()
}

type IntrinsicFunctionNameContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyIntrinsicFunctionNameContext() *IntrinsicFunctionNameContext {
	var p = new(IntrinsicFunctionNameContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_intrinsicFunctionName
	return p
}

func InitEmptyIntrinsicFunctionNameContext(p *IntrinsicFunctionNameContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_intrinsicFunctionName
}

func (*IntrinsicFunctionNameContext) IsIntrinsicFunctionNameContext() {}

func NewIntrinsicFunctionNameContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *IntrinsicFunctionNameContext {
	var p = new(IntrinsicFunctionNameContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_intrinsicFunctionName

	return p
}

func (s *IntrinsicFunctionNameContext) GetParser() antlr.Parser { return s.parser }

func (s *IntrinsicFunctionNameContext) INT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserINT, 0)
}

func (s *IntrinsicFunctionNameContext) FIX() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserFIX, 0)
}

func (s *IntrinsicFunctionNameContext) ABS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserABS, 0)
}

func (s *IntrinsicFunctionNameContext) SGN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSGN, 0)
}

func (s *IntrinsicFunctionNameContext) LEN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLEN, 0)
}

func (s *IntrinsicFunctionNameContext) LENB() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLENB, 0)
}

func (s *IntrinsicFunctionNameContext) CBOOL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCBOOL, 0)
}

func (s *IntrinsicFunctionNameContext) CBYTE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCBYTE, 0)
}

func (s *IntrinsicFunctionNameContext) CCUR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCCUR, 0)
}

func (s *IntrinsicFunctionNameContext) CDATE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCDATE, 0)
}

func (s *IntrinsicFunctionNameContext) CDBL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCDBL, 0)
}

func (s *IntrinsicFunctionNameContext) CINT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCINT, 0)
}

func (s *IntrinsicFunctionNameContext) CLNG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCLNG, 0)
}

func (s *IntrinsicFunctionNameContext) CLNGLNG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCLNGLNG, 0)
}

func (s *IntrinsicFunctionNameContext) CLNGPTR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCLNGPTR, 0)
}

func (s *IntrinsicFunctionNameContext) CSNG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCSNG, 0)
}

func (s *IntrinsicFunctionNameContext) CSTR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCSTR, 0)
}

func (s *IntrinsicFunctionNameContext) CVAR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCVAR, 0)
}

func (s *IntrinsicFunctionNameContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *IntrinsicFunctionNameContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) IntrinsicFunctionName() (localctx IIntrinsicFunctionNameContext) {
	localctx = NewIntrinsicFunctionNameContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 40, VBAConditionalCompilationParserRULE_intrinsicFunctionName)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(459)
		_la = p.GetTokenStream().LA(1)

		if !(((int64(_la) & ^0x3f) == 0 && ((int64(1) << _la) & 34804725234) != 0)) {
			p.GetErrorHandler().RecoverInline(p)
		} else {
			p.GetErrorHandler().ReportMatch(p)
			p.Consume()
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// INameContext is an interface to support dynamic dispatch.
type INameContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	NameValue() INameValueContext
	TypeHint() ITypeHintContext

	// IsNameContext differentiates from other interfaces.
	IsNameContext()
}

type NameContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyNameContext() *NameContext {
	var p = new(NameContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_name
	return p
}

func InitEmptyNameContext(p *NameContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_name
}

func (*NameContext) IsNameContext() {}

func NewNameContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *NameContext {
	var p = new(NameContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_name

	return p
}

func (s *NameContext) GetParser() antlr.Parser { return s.parser }

func (s *NameContext) NameValue() INameValueContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(INameValueContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(INameValueContext)
}

func (s *NameContext) TypeHint() ITypeHintContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(ITypeHintContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(ITypeHintContext)
}

func (s *NameContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *NameContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) Name() (localctx INameContext) {
	localctx = NewNameContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 42, VBAConditionalCompilationParserRULE_name)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(461)
		p.NameValue()
	}
	p.SetState(463)
	p.GetErrorHandler().Sync(p)


	if p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 52, p.GetParserRuleContext()) == 1 {
		{
			p.SetState(462)
			p.TypeHint()
		}

		} else if p.HasError() { // JIM
			goto errorExit
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// INameValueContext is an interface to support dynamic dispatch.
type INameValueContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	IDENTIFIER() antlr.TerminalNode
	Keyword() IKeywordContext
	ForeignName() IForeignNameContext
	StatementKeyword() IStatementKeywordContext
	MarkerKeyword() IMarkerKeywordContext

	// IsNameValueContext differentiates from other interfaces.
	IsNameValueContext()
}

type NameValueContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyNameValueContext() *NameValueContext {
	var p = new(NameValueContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_nameValue
	return p
}

func InitEmptyNameValueContext(p *NameValueContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_nameValue
}

func (*NameValueContext) IsNameValueContext() {}

func NewNameValueContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *NameValueContext {
	var p = new(NameValueContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_nameValue

	return p
}

func (s *NameValueContext) GetParser() antlr.Parser { return s.parser }

func (s *NameValueContext) IDENTIFIER() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserIDENTIFIER, 0)
}

func (s *NameValueContext) Keyword() IKeywordContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IKeywordContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IKeywordContext)
}

func (s *NameValueContext) ForeignName() IForeignNameContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IForeignNameContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IForeignNameContext)
}

func (s *NameValueContext) StatementKeyword() IStatementKeywordContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IStatementKeywordContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IStatementKeywordContext)
}

func (s *NameValueContext) MarkerKeyword() IMarkerKeywordContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IMarkerKeywordContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IMarkerKeywordContext)
}

func (s *NameValueContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *NameValueContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) NameValue() (localctx INameValueContext) {
	localctx = NewNameValueContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 44, VBAConditionalCompilationParserRULE_nameValue)
	p.SetState(470)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}

	switch p.GetTokenStream().LA(1) {
	case VBAConditionalCompilationParserIDENTIFIER:
		p.EnterOuterAlt(localctx, 1)
		{
			p.SetState(465)
			p.Match(VBAConditionalCompilationParserIDENTIFIER)
			if p.HasError() {
					// Recognition error - abort rule
					goto errorExit
			}
		}


	case VBAConditionalCompilationParserABS, VBAConditionalCompilationParserANY, VBAConditionalCompilationParserARRAY, VBAConditionalCompilationParserCBOOL, VBAConditionalCompilationParserCBYTE, VBAConditionalCompilationParserCCUR, VBAConditionalCompilationParserCDATE, VBAConditionalCompilationParserCDBL, VBAConditionalCompilationParserCDEC, VBAConditionalCompilationParserCINT, VBAConditionalCompilationParserCLNG, VBAConditionalCompilationParserCLNGLNG, VBAConditionalCompilationParserCLNGPTR, VBAConditionalCompilationParserCSNG, VBAConditionalCompilationParserCSTR, VBAConditionalCompilationParserCURRENCY, VBAConditionalCompilationParserCVAR, VBAConditionalCompilationParserCVERR, VBAConditionalCompilationParserDEBUG, VBAConditionalCompilationParserDOEVENTS, VBAConditionalCompilationParserFIX, VBAConditionalCompilationParserINPUTB, VBAConditionalCompilationParserINT, VBAConditionalCompilationParserLBOUND, VBAConditionalCompilationParserLEN, VBAConditionalCompilationParserLENB, VBAConditionalCompilationParserLONGLONG, VBAConditionalCompilationParserLONGPTR, VBAConditionalCompilationParserMIDB, VBAConditionalCompilationParserPSET, VBAConditionalCompilationParserSGN, VBAConditionalCompilationParserUBOUND, VBAConditionalCompilationParserACCESS, VBAConditionalCompilationParserADDRESSOF, VBAConditionalCompilationParserALIAS, VBAConditionalCompilationParserAND, VBAConditionalCompilationParserATTRIBUTE, VBAConditionalCompilationParserAPPEND, VBAConditionalCompilationParserBEGIN, VBAConditionalCompilationParserBINARY, VBAConditionalCompilationParserBOOLEAN, VBAConditionalCompilationParserBYVAL, VBAConditionalCompilationParserBYREF, VBAConditionalCompilationParserBYTE, VBAConditionalCompilationParserCLASS, VBAConditionalCompilationParserCLOSE, VBAConditionalCompilationParserDATABASE, VBAConditionalCompilationParserDATE, VBAConditionalCompilationParserDOUBLE, VBAConditionalCompilationParserEND, VBAConditionalCompilationParserEQV, VBAConditionalCompilationParserERROR, VBAConditionalCompilationParserFALSE, VBAConditionalCompilationParserGET, VBAConditionalCompilationParserIMP, VBAConditionalCompilationParserIN, VBAConditionalCompilationParserINPUT, VBAConditionalCompilationParserIS, VBAConditionalCompilationParserINTEGER, VBAConditionalCompilationParserLOCK, VBAConditionalCompilationParserLONG, VBAConditionalCompilationParserLIB, VBAConditionalCompilationParserLIKE, VBAConditionalCompilationParserLINE_INPUT, VBAConditionalCompilationParserLOCK_READ, VBAConditionalCompilationParserLOCK_WRITE, VBAConditionalCompilationParserLOCK_READ_WRITE, VBAConditionalCompilationParserME, VBAConditionalCompilationParserMID, VBAConditionalCompilationParserMOD, VBAConditionalCompilationParserNAME, VBAConditionalCompilationParserNEW, VBAConditionalCompilationParserNOT, VBAConditionalCompilationParserNOTHING, VBAConditionalCompilationParserNULL, VBAConditionalCompilationParserOBJECT, VBAConditionalCompilationParserON_ERROR, VBAConditionalCompilationParserOPEN, VBAConditionalCompilationParserOPTIONAL, VBAConditionalCompilationParserOR, VBAConditionalCompilationParserOUTPUT, VBAConditionalCompilationParserPARAMARRAY, VBAConditionalCompilationParserPRESERVE, VBAConditionalCompilationParserPRINT, VBAConditionalCompilationParserPTRSAFE, VBAConditionalCompilationParserPUT, VBAConditionalCompilationParserRANDOM, VBAConditionalCompilationParserREAD, VBAConditionalCompilationParserREAD_WRITE, VBAConditionalCompilationParserREM, VBAConditionalCompilationParserRESET, VBAConditionalCompilationParserSEEK, VBAConditionalCompilationParserSHARED, VBAConditionalCompilationParserSINGLE, VBAConditionalCompilationParserSPC, VBAConditionalCompilationParserSTEP, VBAConditionalCompilationParserSTRING, VBAConditionalCompilationParserTAB, VBAConditionalCompilationParserTEXT, VBAConditionalCompilationParserTHEN, VBAConditionalCompilationParserTO, VBAConditionalCompilationParserTRUE, VBAConditionalCompilationParserTYPEOF, VBAConditionalCompilationParserUNLOCK, VBAConditionalCompilationParserUNTIL, VBAConditionalCompilationParserVARIANT, VBAConditionalCompilationParserVERSION, VBAConditionalCompilationParserWIDTH, VBAConditionalCompilationParserWITHEVENTS, VBAConditionalCompilationParserWRITE, VBAConditionalCompilationParserXOR, VBAConditionalCompilationParserLOAD, VBAConditionalCompilationParserMIDBTYPESUFFIX, VBAConditionalCompilationParserMIDTYPESUFFIX, VBAConditionalCompilationParserRESUME_NEXT:
		p.EnterOuterAlt(localctx, 2)
		{
			p.SetState(466)
			p.Keyword()
		}


	case VBAConditionalCompilationParserL_SQUARE_BRACKET:
		p.EnterOuterAlt(localctx, 3)
		{
			p.SetState(467)
			p.ForeignName()
		}


	case VBAConditionalCompilationParserEXIT, VBAConditionalCompilationParserOPTION, VBAConditionalCompilationParserCALL, VBAConditionalCompilationParserCASE, VBAConditionalCompilationParserCONST, VBAConditionalCompilationParserDECLARE, VBAConditionalCompilationParserDEFBOOL, VBAConditionalCompilationParserDEFBYTE, VBAConditionalCompilationParserDEFDATE, VBAConditionalCompilationParserDEFDBL, VBAConditionalCompilationParserDEFCUR, VBAConditionalCompilationParserDEFINT, VBAConditionalCompilationParserDEFLNG, VBAConditionalCompilationParserDEFLNGLNG, VBAConditionalCompilationParserDEFLNGPTR, VBAConditionalCompilationParserDEFOBJ, VBAConditionalCompilationParserDEFSNG, VBAConditionalCompilationParserDEFSTR, VBAConditionalCompilationParserDEFVAR, VBAConditionalCompilationParserDIM, VBAConditionalCompilationParserDO, VBAConditionalCompilationParserELSE, VBAConditionalCompilationParserELSEIF, VBAConditionalCompilationParserEND_SELECT, VBAConditionalCompilationParserEND_WITH, VBAConditionalCompilationParserENUM, VBAConditionalCompilationParserERASE, VBAConditionalCompilationParserEVENT, VBAConditionalCompilationParserEXIT_DO, VBAConditionalCompilationParserEXIT_FOR, VBAConditionalCompilationParserEXIT_FUNCTION, VBAConditionalCompilationParserEXIT_PROPERTY, VBAConditionalCompilationParserEXIT_SUB, VBAConditionalCompilationParserFRIEND, VBAConditionalCompilationParserFOR, VBAConditionalCompilationParserFUNCTION, VBAConditionalCompilationParserGLOBAL, VBAConditionalCompilationParserGOSUB, VBAConditionalCompilationParserGOTO, VBAConditionalCompilationParserIF, VBAConditionalCompilationParserIMPLEMENTS, VBAConditionalCompilationParserLOOP, VBAConditionalCompilationParserLET, VBAConditionalCompilationParserLSET, VBAConditionalCompilationParserNEXT, VBAConditionalCompilationParserON, VBAConditionalCompilationParserPRIVATE, VBAConditionalCompilationParserPUBLIC, VBAConditionalCompilationParserRAISEEVENT, VBAConditionalCompilationParserREDIM, VBAConditionalCompilationParserRESUME, VBAConditionalCompilationParserRETURN, VBAConditionalCompilationParserRSET, VBAConditionalCompilationParserSELECT, VBAConditionalCompilationParserSET, VBAConditionalCompilationParserSTATIC, VBAConditionalCompilationParserSTOP, VBAConditionalCompilationParserSUB, VBAConditionalCompilationParserTYPE, VBAConditionalCompilationParserWEND, VBAConditionalCompilationParserWHILE, VBAConditionalCompilationParserWITH:
		p.EnterOuterAlt(localctx, 4)
		{
			p.SetState(468)
			p.StatementKeyword()
		}


	case VBAConditionalCompilationParserAS:
		p.EnterOuterAlt(localctx, 5)
		{
			p.SetState(469)
			p.MarkerKeyword()
		}



	default:
		p.SetError(antlr.NewNoViableAltException(p, nil, nil, nil, nil, nil))
		goto errorExit
	}


errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IForeignNameContext is an interface to support dynamic dispatch.
type IForeignNameContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	L_SQUARE_BRACKET() antlr.TerminalNode
	R_SQUARE_BRACKET() antlr.TerminalNode
	AllForeignIdentifier() []IForeignIdentifierContext
	ForeignIdentifier(i int) IForeignIdentifierContext

	// IsForeignNameContext differentiates from other interfaces.
	IsForeignNameContext()
}

type ForeignNameContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyForeignNameContext() *ForeignNameContext {
	var p = new(ForeignNameContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_foreignName
	return p
}

func InitEmptyForeignNameContext(p *ForeignNameContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_foreignName
}

func (*ForeignNameContext) IsForeignNameContext() {}

func NewForeignNameContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *ForeignNameContext {
	var p = new(ForeignNameContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_foreignName

	return p
}

func (s *ForeignNameContext) GetParser() antlr.Parser { return s.parser }

func (s *ForeignNameContext) L_SQUARE_BRACKET() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserL_SQUARE_BRACKET, 0)
}

func (s *ForeignNameContext) R_SQUARE_BRACKET() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserR_SQUARE_BRACKET, 0)
}

func (s *ForeignNameContext) AllForeignIdentifier() []IForeignIdentifierContext {
	children := s.GetChildren()
	len := 0
	for _, ctx := range children {
		if _, ok := ctx.(IForeignIdentifierContext); ok {
			len++
		}
	}

	tst := make([]IForeignIdentifierContext, len)
	i := 0
	for _, ctx := range children {
		if t, ok := ctx.(IForeignIdentifierContext); ok {
			tst[i] = t.(IForeignIdentifierContext)
			i++
		}
	}

	return tst
}

func (s *ForeignNameContext) ForeignIdentifier(i int) IForeignIdentifierContext {
	var t antlr.RuleContext;
	j := 0
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IForeignIdentifierContext); ok {
			if j == i {
				t = ctx.(antlr.RuleContext);
				break
			}
			j++
		}
	}

	if t == nil {
		return nil
	}

	return t.(IForeignIdentifierContext)
}

func (s *ForeignNameContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *ForeignNameContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) ForeignName() (localctx IForeignNameContext) {
	localctx = NewForeignNameContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 46, VBAConditionalCompilationParserRULE_foreignName)
	var _alt int

	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(472)
		p.Match(VBAConditionalCompilationParserL_SQUARE_BRACKET)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	p.SetState(476)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_alt = p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 54, p.GetParserRuleContext())
	if p.HasError() {
		goto errorExit
	}
	for _alt != 2 && _alt != antlr.ATNInvalidAltNumber {
		if _alt == 1 {
			{
				p.SetState(473)
				p.ForeignIdentifier()
			}


		}
		p.SetState(478)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_alt = p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 54, p.GetParserRuleContext())
		if p.HasError() {
			goto errorExit
		}
	}
	{
		p.SetState(479)
		p.Match(VBAConditionalCompilationParserR_SQUARE_BRACKET)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IForeignIdentifierContext is an interface to support dynamic dispatch.
type IForeignIdentifierContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	L_SQUARE_BRACKET() antlr.TerminalNode
	ForeignName() IForeignNameContext

	// IsForeignIdentifierContext differentiates from other interfaces.
	IsForeignIdentifierContext()
}

type ForeignIdentifierContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyForeignIdentifierContext() *ForeignIdentifierContext {
	var p = new(ForeignIdentifierContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_foreignIdentifier
	return p
}

func InitEmptyForeignIdentifierContext(p *ForeignIdentifierContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_foreignIdentifier
}

func (*ForeignIdentifierContext) IsForeignIdentifierContext() {}

func NewForeignIdentifierContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *ForeignIdentifierContext {
	var p = new(ForeignIdentifierContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_foreignIdentifier

	return p
}

func (s *ForeignIdentifierContext) GetParser() antlr.Parser { return s.parser }

func (s *ForeignIdentifierContext) L_SQUARE_BRACKET() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserL_SQUARE_BRACKET, 0)
}

func (s *ForeignIdentifierContext) ForeignName() IForeignNameContext {
	var t antlr.RuleContext;
	for _, ctx := range s.GetChildren() {
		if _, ok := ctx.(IForeignNameContext); ok {
			t = ctx.(antlr.RuleContext);
			break
		}
	}

	if t == nil {
		return nil
	}

	return t.(IForeignNameContext)
}

func (s *ForeignIdentifierContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *ForeignIdentifierContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) ForeignIdentifier() (localctx IForeignIdentifierContext) {
	localctx = NewForeignIdentifierContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 48, VBAConditionalCompilationParserRULE_foreignIdentifier)
	var _la int

	p.SetState(483)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}

	switch p.GetTokenStream().LA(1) {
	case VBAConditionalCompilationParserABS, VBAConditionalCompilationParserANY, VBAConditionalCompilationParserARRAY, VBAConditionalCompilationParserCBOOL, VBAConditionalCompilationParserCBYTE, VBAConditionalCompilationParserCCUR, VBAConditionalCompilationParserCDATE, VBAConditionalCompilationParserCDBL, VBAConditionalCompilationParserCDEC, VBAConditionalCompilationParserCINT, VBAConditionalCompilationParserCIRCLE, VBAConditionalCompilationParserCLNG, VBAConditionalCompilationParserCLNGLNG, VBAConditionalCompilationParserCLNGPTR, VBAConditionalCompilationParserCSNG, VBAConditionalCompilationParserCSTR, VBAConditionalCompilationParserCURRENCY, VBAConditionalCompilationParserCVAR, VBAConditionalCompilationParserCVERR, VBAConditionalCompilationParserDEBUG, VBAConditionalCompilationParserDOEVENTS, VBAConditionalCompilationParserEXIT, VBAConditionalCompilationParserFIX, VBAConditionalCompilationParserINPUTB, VBAConditionalCompilationParserINT, VBAConditionalCompilationParserLBOUND, VBAConditionalCompilationParserLEN, VBAConditionalCompilationParserLENB, VBAConditionalCompilationParserLONGLONG, VBAConditionalCompilationParserLONGPTR, VBAConditionalCompilationParserMIDB, VBAConditionalCompilationParserOPTION, VBAConditionalCompilationParserPSET, VBAConditionalCompilationParserSCALE, VBAConditionalCompilationParserSGN, VBAConditionalCompilationParserUBOUND, VBAConditionalCompilationParserCOMMA, VBAConditionalCompilationParserCOLON, VBAConditionalCompilationParserSEMICOLON, VBAConditionalCompilationParserEXCLAMATIONPOINT, VBAConditionalCompilationParserDOT, VBAConditionalCompilationParserHASH, VBAConditionalCompilationParserAT, VBAConditionalCompilationParserPERCENT, VBAConditionalCompilationParserDOLLAR, VBAConditionalCompilationParserAMPERSAND, VBAConditionalCompilationParserACCESS, VBAConditionalCompilationParserADDRESSOF, VBAConditionalCompilationParserALIAS, VBAConditionalCompilationParserAND, VBAConditionalCompilationParserATTRIBUTE, VBAConditionalCompilationParserAPPEND, VBAConditionalCompilationParserAS, VBAConditionalCompilationParserBEGINPROPERTY, VBAConditionalCompilationParserBEGIN, VBAConditionalCompilationParserBINARY, VBAConditionalCompilationParserBOOLEAN, VBAConditionalCompilationParserBYVAL, VBAConditionalCompilationParserBYREF, VBAConditionalCompilationParserBYTE, VBAConditionalCompilationParserCALL, VBAConditionalCompilationParserCASE, VBAConditionalCompilationParserCDECL, VBAConditionalCompilationParserCLASS, VBAConditionalCompilationParserCLOSE, VBAConditionalCompilationParserCONST, VBAConditionalCompilationParserDATABASE, VBAConditionalCompilationParserDATE, VBAConditionalCompilationParserDECLARE, VBAConditionalCompilationParserDEFBOOL, VBAConditionalCompilationParserDEFBYTE, VBAConditionalCompilationParserDEFDATE, VBAConditionalCompilationParserDEFDBL, VBAConditionalCompilationParserDEFCUR, VBAConditionalCompilationParserDEFINT, VBAConditionalCompilationParserDEFLNG, VBAConditionalCompilationParserDEFLNGLNG, VBAConditionalCompilationParserDEFLNGPTR, VBAConditionalCompilationParserDEFOBJ, VBAConditionalCompilationParserDEFSNG, VBAConditionalCompilationParserDEFSTR, VBAConditionalCompilationParserDEFVAR, VBAConditionalCompilationParserDIM, VBAConditionalCompilationParserDO, VBAConditionalCompilationParserDOUBLE, VBAConditionalCompilationParserEACH, VBAConditionalCompilationParserELSE, VBAConditionalCompilationParserELSEIF, VBAConditionalCompilationParserEMPTY, VBAConditionalCompilationParserEND_ENUM, VBAConditionalCompilationParserEND_FUNCTION, VBAConditionalCompilationParserEND_IF, VBAConditionalCompilationParserENDPROPERTY, VBAConditionalCompilationParserEND_PROPERTY, VBAConditionalCompilationParserEND_SELECT, VBAConditionalCompilationParserEND_SUB, VBAConditionalCompilationParserEND_TYPE, VBAConditionalCompilationParserEND_WITH, VBAConditionalCompilationParserEND, VBAConditionalCompilationParserENUM, VBAConditionalCompilationParserEQV, VBAConditionalCompilationParserERASE, VBAConditionalCompilationParserERROR, VBAConditionalCompilationParserEVENT, VBAConditionalCompilationParserEXIT_DO, VBAConditionalCompilationParserEXIT_FOR, VBAConditionalCompilationParserEXIT_FUNCTION, VBAConditionalCompilationParserEXIT_PROPERTY, VBAConditionalCompilationParserEXIT_SUB, VBAConditionalCompilationParserFALSE, VBAConditionalCompilationParserFRIEND, VBAConditionalCompilationParserFOR, VBAConditionalCompilationParserFUNCTION, VBAConditionalCompilationParserGET, VBAConditionalCompilationParserGLOBAL, VBAConditionalCompilationParserGOSUB, VBAConditionalCompilationParserGOTO, VBAConditionalCompilationParserIF, VBAConditionalCompilationParserIMP, VBAConditionalCompilationParserIMPLEMENTS, VBAConditionalCompilationParserIN, VBAConditionalCompilationParserINPUT, VBAConditionalCompilationParserIS, VBAConditionalCompilationParserINTEGER, VBAConditionalCompilationParserLOCK, VBAConditionalCompilationParserLONG, VBAConditionalCompilationParserLOOP, VBAConditionalCompilationParserLET, VBAConditionalCompilationParserLIB, VBAConditionalCompilationParserLIKE, VBAConditionalCompilationParserLINE_INPUT, VBAConditionalCompilationParserLOCK_READ, VBAConditionalCompilationParserLOCK_WRITE, VBAConditionalCompilationParserLOCK_READ_WRITE, VBAConditionalCompilationParserLSET, VBAConditionalCompilationParserME, VBAConditionalCompilationParserMID, VBAConditionalCompilationParserMOD, VBAConditionalCompilationParserNAME, VBAConditionalCompilationParserNEXT, VBAConditionalCompilationParserNEW, VBAConditionalCompilationParserNOT, VBAConditionalCompilationParserNOTHING, VBAConditionalCompilationParserNULL, VBAConditionalCompilationParserOBJECT, VBAConditionalCompilationParserON, VBAConditionalCompilationParserON_ERROR, VBAConditionalCompilationParserON_LOCAL_ERROR, VBAConditionalCompilationParserOPEN, VBAConditionalCompilationParserOPTIONAL, VBAConditionalCompilationParserOPTION_BASE, VBAConditionalCompilationParserOPTION_EXPLICIT, VBAConditionalCompilationParserOPTION_COMPARE, VBAConditionalCompilationParserOPTION_PRIVATE_MODULE, VBAConditionalCompilationParserOR, VBAConditionalCompilationParserOUTPUT, VBAConditionalCompilationParserPARAMARRAY, VBAConditionalCompilationParserPRESERVE, VBAConditionalCompilationParserPRINT, VBAConditionalCompilationParserPRIVATE, VBAConditionalCompilationParserPROPERTY_GET, VBAConditionalCompilationParserPROPERTY_LET, VBAConditionalCompilationParserPROPERTY_SET, VBAConditionalCompilationParserPTRSAFE, VBAConditionalCompilationParserPUBLIC, VBAConditionalCompilationParserPUT, VBAConditionalCompilationParserRANDOM, VBAConditionalCompilationParserRANDOMIZE, VBAConditionalCompilationParserRAISEEVENT, VBAConditionalCompilationParserREAD, VBAConditionalCompilationParserREAD_WRITE, VBAConditionalCompilationParserREDIM, VBAConditionalCompilationParserREM, VBAConditionalCompilationParserRESET, VBAConditionalCompilationParserRESUME, VBAConditionalCompilationParserRETURN, VBAConditionalCompilationParserRSET, VBAConditionalCompilationParserSEEK, VBAConditionalCompilationParserSELECT, VBAConditionalCompilationParserSET, VBAConditionalCompilationParserSHARED, VBAConditionalCompilationParserSINGLE, VBAConditionalCompilationParserSPC, VBAConditionalCompilationParserSTATIC, VBAConditionalCompilationParserSTEP, VBAConditionalCompilationParserSTOP, VBAConditionalCompilationParserSTRING, VBAConditionalCompilationParserSUB, VBAConditionalCompilationParserTAB, VBAConditionalCompilationParserTEXT, VBAConditionalCompilationParserTHEN, VBAConditionalCompilationParserTO, VBAConditionalCompilationParserTRUE, VBAConditionalCompilationParserTYPE, VBAConditionalCompilationParserTYPEOF, VBAConditionalCompilationParserUNLOCK, VBAConditionalCompilationParserUNTIL, VBAConditionalCompilationParserVARIANT, VBAConditionalCompilationParserVERSION, VBAConditionalCompilationParserWEND, VBAConditionalCompilationParserWHILE, VBAConditionalCompilationParserWIDTH, VBAConditionalCompilationParserWITH, VBAConditionalCompilationParserWITHEVENTS, VBAConditionalCompilationParserWRITE, VBAConditionalCompilationParserXOR, VBAConditionalCompilationParserASSIGN, VBAConditionalCompilationParserDIV, VBAConditionalCompilationParserINTDIV, VBAConditionalCompilationParserEQ, VBAConditionalCompilationParserGEQ, VBAConditionalCompilationParserGT, VBAConditionalCompilationParserLEQ, VBAConditionalCompilationParserLPAREN, VBAConditionalCompilationParserLT, VBAConditionalCompilationParserMINUS, VBAConditionalCompilationParserMULT, VBAConditionalCompilationParserNEQ, VBAConditionalCompilationParserPLUS, VBAConditionalCompilationParserPOW, VBAConditionalCompilationParserRPAREN, VBAConditionalCompilationParserR_SQUARE_BRACKET, VBAConditionalCompilationParserL_BRACE, VBAConditionalCompilationParserR_BRACE, VBAConditionalCompilationParserSTRINGLITERAL, VBAConditionalCompilationParserOCTLITERAL, VBAConditionalCompilationParserHEXLITERAL, VBAConditionalCompilationParserFLOATLITERAL, VBAConditionalCompilationParserINTEGERLITERAL, VBAConditionalCompilationParserDATELITERAL, VBAConditionalCompilationParserNEWLINE, VBAConditionalCompilationParserSINGLEQUOTE, VBAConditionalCompilationParserUNDERSCORE, VBAConditionalCompilationParserWS, VBAConditionalCompilationParserGUIDLITERAL, VBAConditionalCompilationParserIDENTIFIER, VBAConditionalCompilationParserLINE_CONTINUATION, VBAConditionalCompilationParserBARE_HEX_LITERAL, VBAConditionalCompilationParserERRORCHAR, VBAConditionalCompilationParserLOAD, VBAConditionalCompilationParserMIDBTYPESUFFIX, VBAConditionalCompilationParserMIDTYPESUFFIX, VBAConditionalCompilationParserRESUME_NEXT:
		p.EnterOuterAlt(localctx, 1)
		{
			p.SetState(481)
			_la = p.GetTokenStream().LA(1)

			if _la <= 0 || _la == VBAConditionalCompilationParserL_SQUARE_BRACKET  {
				p.GetErrorHandler().RecoverInline(p)
			} else {
				p.GetErrorHandler().ReportMatch(p)
				p.Consume()
			}
		}


	case VBAConditionalCompilationParserL_SQUARE_BRACKET:
		p.EnterOuterAlt(localctx, 2)
		{
			p.SetState(482)
			p.ForeignName()
		}



	default:
		p.SetError(antlr.NewNoViableAltException(p, nil, nil, nil, nil, nil))
		goto errorExit
	}


errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ITypeHintContext is an interface to support dynamic dispatch.
type ITypeHintContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	PERCENT() antlr.TerminalNode
	AMPERSAND() antlr.TerminalNode
	POW() antlr.TerminalNode
	EXCLAMATIONPOINT() antlr.TerminalNode
	HASH() antlr.TerminalNode
	AT() antlr.TerminalNode
	DOLLAR() antlr.TerminalNode

	// IsTypeHintContext differentiates from other interfaces.
	IsTypeHintContext()
}

type TypeHintContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyTypeHintContext() *TypeHintContext {
	var p = new(TypeHintContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_typeHint
	return p
}

func InitEmptyTypeHintContext(p *TypeHintContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_typeHint
}

func (*TypeHintContext) IsTypeHintContext() {}

func NewTypeHintContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *TypeHintContext {
	var p = new(TypeHintContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_typeHint

	return p
}

func (s *TypeHintContext) GetParser() antlr.Parser { return s.parser }

func (s *TypeHintContext) PERCENT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPERCENT, 0)
}

func (s *TypeHintContext) AMPERSAND() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserAMPERSAND, 0)
}

func (s *TypeHintContext) POW() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPOW, 0)
}

func (s *TypeHintContext) EXCLAMATIONPOINT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEXCLAMATIONPOINT, 0)
}

func (s *TypeHintContext) HASH() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserHASH, 0)
}

func (s *TypeHintContext) AT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserAT, 0)
}

func (s *TypeHintContext) DOLLAR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDOLLAR, 0)
}

func (s *TypeHintContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *TypeHintContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) TypeHint() (localctx ITypeHintContext) {
	localctx = NewTypeHintContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 50, VBAConditionalCompilationParserRULE_typeHint)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(485)
		_la = p.GetTokenStream().LA(1)

		if !(((int64(_la) & ^0x3f) == 0 && ((int64(1) << _la) & 137438953472000) != 0) || _la == VBAConditionalCompilationParserPOW) {
			p.GetErrorHandler().RecoverInline(p)
		} else {
			p.GetErrorHandler().ReportMatch(p)
			p.Consume()
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ILiteralContext is an interface to support dynamic dispatch.
type ILiteralContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	DATELITERAL() antlr.TerminalNode
	HEXLITERAL() antlr.TerminalNode
	OCTLITERAL() antlr.TerminalNode
	FLOATLITERAL() antlr.TerminalNode
	INTEGERLITERAL() antlr.TerminalNode
	STRINGLITERAL() antlr.TerminalNode
	TRUE() antlr.TerminalNode
	FALSE() antlr.TerminalNode
	NOTHING() antlr.TerminalNode
	NULL() antlr.TerminalNode
	EMPTY() antlr.TerminalNode

	// IsLiteralContext differentiates from other interfaces.
	IsLiteralContext()
}

type LiteralContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyLiteralContext() *LiteralContext {
	var p = new(LiteralContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_literal
	return p
}

func InitEmptyLiteralContext(p *LiteralContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_literal
}

func (*LiteralContext) IsLiteralContext() {}

func NewLiteralContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *LiteralContext {
	var p = new(LiteralContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_literal

	return p
}

func (s *LiteralContext) GetParser() antlr.Parser { return s.parser }

func (s *LiteralContext) DATELITERAL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDATELITERAL, 0)
}

func (s *LiteralContext) HEXLITERAL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserHEXLITERAL, 0)
}

func (s *LiteralContext) OCTLITERAL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserOCTLITERAL, 0)
}

func (s *LiteralContext) FLOATLITERAL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserFLOATLITERAL, 0)
}

func (s *LiteralContext) INTEGERLITERAL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserINTEGERLITERAL, 0)
}

func (s *LiteralContext) STRINGLITERAL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSTRINGLITERAL, 0)
}

func (s *LiteralContext) TRUE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserTRUE, 0)
}

func (s *LiteralContext) FALSE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserFALSE, 0)
}

func (s *LiteralContext) NOTHING() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNOTHING, 0)
}

func (s *LiteralContext) NULL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNULL, 0)
}

func (s *LiteralContext) EMPTY() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEMPTY, 0)
}

func (s *LiteralContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *LiteralContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) Literal() (localctx ILiteralContext) {
	localctx = NewLiteralContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 52, VBAConditionalCompilationParserRULE_literal)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(487)
		_la = p.GetTokenStream().LA(1)

		if !(((int64((_la - 89)) & ^0x3f) == 0 && ((int64(1) << (_la - 89)) & 54043195530543105) != 0) || ((int64((_la - 193)) & ^0x3f) == 0 && ((int64(1) << (_la - 193)) & 541165879297) != 0)) {
			p.GetErrorHandler().RecoverInline(p)
		} else {
			p.GetErrorHandler().ReportMatch(p)
			p.Consume()
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// ICommentContext is an interface to support dynamic dispatch.
type ICommentContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	SINGLEQUOTE() antlr.TerminalNode
	AllLINE_CONTINUATION() []antlr.TerminalNode
	LINE_CONTINUATION(i int) antlr.TerminalNode
	AllNEWLINE() []antlr.TerminalNode
	NEWLINE(i int) antlr.TerminalNode

	// IsCommentContext differentiates from other interfaces.
	IsCommentContext()
}

type CommentContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyCommentContext() *CommentContext {
	var p = new(CommentContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_comment
	return p
}

func InitEmptyCommentContext(p *CommentContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_comment
}

func (*CommentContext) IsCommentContext() {}

func NewCommentContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *CommentContext {
	var p = new(CommentContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_comment

	return p
}

func (s *CommentContext) GetParser() antlr.Parser { return s.parser }

func (s *CommentContext) SINGLEQUOTE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSINGLEQUOTE, 0)
}

func (s *CommentContext) AllLINE_CONTINUATION() []antlr.TerminalNode {
	return s.GetTokens(VBAConditionalCompilationParserLINE_CONTINUATION)
}

func (s *CommentContext) LINE_CONTINUATION(i int) antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLINE_CONTINUATION, i)
}

func (s *CommentContext) AllNEWLINE() []antlr.TerminalNode {
	return s.GetTokens(VBAConditionalCompilationParserNEWLINE)
}

func (s *CommentContext) NEWLINE(i int) antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNEWLINE, i)
}

func (s *CommentContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *CommentContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) Comment() (localctx ICommentContext) {
	localctx = NewCommentContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 54, VBAConditionalCompilationParserRULE_comment)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(489)
		p.Match(VBAConditionalCompilationParserSINGLEQUOTE)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}
	p.SetState(494)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_la = p.GetTokenStream().LA(1)


	for ((int64(_la) & ^0x3f) == 0 && ((int64(1) << _la) & -2) != 0) || ((int64((_la - 64)) & ^0x3f) == 0 && ((int64(1) << (_la - 64)) & -1) != 0) || ((int64((_la - 128)) & ^0x3f) == 0 && ((int64(1) << (_la - 128)) & -1) != 0) || ((int64((_la - 192)) & ^0x3f) == 0 && ((int64(1) << (_la - 192)) & 9006099743113215) != 0) {
		p.SetState(492)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
			goto errorExit
		}

		switch p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 56, p.GetParserRuleContext()) {
		case 1:
			{
				p.SetState(490)
				p.Match(VBAConditionalCompilationParserLINE_CONTINUATION)
				if p.HasError() {
						// Recognition error - abort rule
						goto errorExit
				}
			}


		case 2:
			{
				p.SetState(491)
				_la = p.GetTokenStream().LA(1)

				if _la <= 0 || _la == VBAConditionalCompilationParserNEWLINE  {
					p.GetErrorHandler().RecoverInline(p)
				} else {
					p.GetErrorHandler().ReportMatch(p)
					p.Consume()
				}
			}

		case antlr.ATNInvalidAltNumber:
			goto errorExit
		}

		p.SetState(496)
		p.GetErrorHandler().Sync(p)
		if p.HasError() {
	    	goto errorExit
	    }
		_la = p.GetTokenStream().LA(1)
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IKeywordContext is an interface to support dynamic dispatch.
type IKeywordContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	ABS() antlr.TerminalNode
	ADDRESSOF() antlr.TerminalNode
	ALIAS() antlr.TerminalNode
	AND() antlr.TerminalNode
	ANY() antlr.TerminalNode
	ARRAY() antlr.TerminalNode
	ATTRIBUTE() antlr.TerminalNode
	BEGIN() antlr.TerminalNode
	BOOLEAN() antlr.TerminalNode
	BYREF() antlr.TerminalNode
	BYTE() antlr.TerminalNode
	BYVAL() antlr.TerminalNode
	CBOOL() antlr.TerminalNode
	CBYTE() antlr.TerminalNode
	CCUR() antlr.TerminalNode
	CDATE() antlr.TerminalNode
	CDBL() antlr.TerminalNode
	CDEC() antlr.TerminalNode
	CINT() antlr.TerminalNode
	CLASS() antlr.TerminalNode
	CLNG() antlr.TerminalNode
	CLNGLNG() antlr.TerminalNode
	CLNGPTR() antlr.TerminalNode
	CSNG() antlr.TerminalNode
	CSTR() antlr.TerminalNode
	CURRENCY() antlr.TerminalNode
	CVAR() antlr.TerminalNode
	CVERR() antlr.TerminalNode
	DATABASE() antlr.TerminalNode
	DATE() antlr.TerminalNode
	DEBUG() antlr.TerminalNode
	DOEVENTS() antlr.TerminalNode
	DOUBLE() antlr.TerminalNode
	END() antlr.TerminalNode
	EQV() antlr.TerminalNode
	FALSE() antlr.TerminalNode
	FIX() antlr.TerminalNode
	IMP() antlr.TerminalNode
	IN() antlr.TerminalNode
	INPUTB() antlr.TerminalNode
	INT() antlr.TerminalNode
	INTEGER() antlr.TerminalNode
	IS() antlr.TerminalNode
	LBOUND() antlr.TerminalNode
	LEN() antlr.TerminalNode
	LENB() antlr.TerminalNode
	LIB() antlr.TerminalNode
	LIKE() antlr.TerminalNode
	LOAD() antlr.TerminalNode
	LONG() antlr.TerminalNode
	LONGLONG() antlr.TerminalNode
	LONGPTR() antlr.TerminalNode
	ME() antlr.TerminalNode
	MID() antlr.TerminalNode
	MIDB() antlr.TerminalNode
	MIDBTYPESUFFIX() antlr.TerminalNode
	MIDTYPESUFFIX() antlr.TerminalNode
	MOD() antlr.TerminalNode
	NEW() antlr.TerminalNode
	NOT() antlr.TerminalNode
	NOTHING() antlr.TerminalNode
	NULL() antlr.TerminalNode
	OBJECT() antlr.TerminalNode
	OPTIONAL() antlr.TerminalNode
	OR() antlr.TerminalNode
	PARAMARRAY() antlr.TerminalNode
	PRESERVE() antlr.TerminalNode
	PSET() antlr.TerminalNode
	PTRSAFE() antlr.TerminalNode
	REM() antlr.TerminalNode
	SGN() antlr.TerminalNode
	SINGLE() antlr.TerminalNode
	SPC() antlr.TerminalNode
	STRING() antlr.TerminalNode
	TAB() antlr.TerminalNode
	TEXT() antlr.TerminalNode
	THEN() antlr.TerminalNode
	TO() antlr.TerminalNode
	TRUE() antlr.TerminalNode
	TYPEOF() antlr.TerminalNode
	UBOUND() antlr.TerminalNode
	UNTIL() antlr.TerminalNode
	VARIANT() antlr.TerminalNode
	VERSION() antlr.TerminalNode
	WITHEVENTS() antlr.TerminalNode
	XOR() antlr.TerminalNode
	STEP() antlr.TerminalNode
	ON_ERROR() antlr.TerminalNode
	RESUME_NEXT() antlr.TerminalNode
	ERROR() antlr.TerminalNode
	APPEND() antlr.TerminalNode
	BINARY() antlr.TerminalNode
	OUTPUT() antlr.TerminalNode
	RANDOM() antlr.TerminalNode
	ACCESS() antlr.TerminalNode
	READ() antlr.TerminalNode
	WRITE() antlr.TerminalNode
	READ_WRITE() antlr.TerminalNode
	SHARED() antlr.TerminalNode
	LOCK_READ() antlr.TerminalNode
	LOCK_WRITE() antlr.TerminalNode
	LOCK_READ_WRITE() antlr.TerminalNode
	LINE_INPUT() antlr.TerminalNode
	RESET() antlr.TerminalNode
	WIDTH() antlr.TerminalNode
	PRINT() antlr.TerminalNode
	GET() antlr.TerminalNode
	PUT() antlr.TerminalNode
	CLOSE() antlr.TerminalNode
	INPUT() antlr.TerminalNode
	LOCK() antlr.TerminalNode
	OPEN() antlr.TerminalNode
	SEEK() antlr.TerminalNode
	UNLOCK() antlr.TerminalNode
	NAME() antlr.TerminalNode

	// IsKeywordContext differentiates from other interfaces.
	IsKeywordContext()
}

type KeywordContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyKeywordContext() *KeywordContext {
	var p = new(KeywordContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_keyword
	return p
}

func InitEmptyKeywordContext(p *KeywordContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_keyword
}

func (*KeywordContext) IsKeywordContext() {}

func NewKeywordContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *KeywordContext {
	var p = new(KeywordContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_keyword

	return p
}

func (s *KeywordContext) GetParser() antlr.Parser { return s.parser }

func (s *KeywordContext) ABS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserABS, 0)
}

func (s *KeywordContext) ADDRESSOF() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserADDRESSOF, 0)
}

func (s *KeywordContext) ALIAS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserALIAS, 0)
}

func (s *KeywordContext) AND() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserAND, 0)
}

func (s *KeywordContext) ANY() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserANY, 0)
}

func (s *KeywordContext) ARRAY() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserARRAY, 0)
}

func (s *KeywordContext) ATTRIBUTE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserATTRIBUTE, 0)
}

func (s *KeywordContext) BEGIN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserBEGIN, 0)
}

func (s *KeywordContext) BOOLEAN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserBOOLEAN, 0)
}

func (s *KeywordContext) BYREF() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserBYREF, 0)
}

func (s *KeywordContext) BYTE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserBYTE, 0)
}

func (s *KeywordContext) BYVAL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserBYVAL, 0)
}

func (s *KeywordContext) CBOOL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCBOOL, 0)
}

func (s *KeywordContext) CBYTE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCBYTE, 0)
}

func (s *KeywordContext) CCUR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCCUR, 0)
}

func (s *KeywordContext) CDATE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCDATE, 0)
}

func (s *KeywordContext) CDBL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCDBL, 0)
}

func (s *KeywordContext) CDEC() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCDEC, 0)
}

func (s *KeywordContext) CINT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCINT, 0)
}

func (s *KeywordContext) CLASS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCLASS, 0)
}

func (s *KeywordContext) CLNG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCLNG, 0)
}

func (s *KeywordContext) CLNGLNG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCLNGLNG, 0)
}

func (s *KeywordContext) CLNGPTR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCLNGPTR, 0)
}

func (s *KeywordContext) CSNG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCSNG, 0)
}

func (s *KeywordContext) CSTR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCSTR, 0)
}

func (s *KeywordContext) CURRENCY() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCURRENCY, 0)
}

func (s *KeywordContext) CVAR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCVAR, 0)
}

func (s *KeywordContext) CVERR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCVERR, 0)
}

func (s *KeywordContext) DATABASE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDATABASE, 0)
}

func (s *KeywordContext) DATE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDATE, 0)
}

func (s *KeywordContext) DEBUG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEBUG, 0)
}

func (s *KeywordContext) DOEVENTS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDOEVENTS, 0)
}

func (s *KeywordContext) DOUBLE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDOUBLE, 0)
}

func (s *KeywordContext) END() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEND, 0)
}

func (s *KeywordContext) EQV() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEQV, 0)
}

func (s *KeywordContext) FALSE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserFALSE, 0)
}

func (s *KeywordContext) FIX() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserFIX, 0)
}

func (s *KeywordContext) IMP() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserIMP, 0)
}

func (s *KeywordContext) IN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserIN, 0)
}

func (s *KeywordContext) INPUTB() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserINPUTB, 0)
}

func (s *KeywordContext) INT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserINT, 0)
}

func (s *KeywordContext) INTEGER() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserINTEGER, 0)
}

func (s *KeywordContext) IS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserIS, 0)
}

func (s *KeywordContext) LBOUND() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLBOUND, 0)
}

func (s *KeywordContext) LEN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLEN, 0)
}

func (s *KeywordContext) LENB() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLENB, 0)
}

func (s *KeywordContext) LIB() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLIB, 0)
}

func (s *KeywordContext) LIKE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLIKE, 0)
}

func (s *KeywordContext) LOAD() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLOAD, 0)
}

func (s *KeywordContext) LONG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLONG, 0)
}

func (s *KeywordContext) LONGLONG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLONGLONG, 0)
}

func (s *KeywordContext) LONGPTR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLONGPTR, 0)
}

func (s *KeywordContext) ME() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserME, 0)
}

func (s *KeywordContext) MID() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserMID, 0)
}

func (s *KeywordContext) MIDB() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserMIDB, 0)
}

func (s *KeywordContext) MIDBTYPESUFFIX() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserMIDBTYPESUFFIX, 0)
}

func (s *KeywordContext) MIDTYPESUFFIX() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserMIDTYPESUFFIX, 0)
}

func (s *KeywordContext) MOD() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserMOD, 0)
}

func (s *KeywordContext) NEW() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNEW, 0)
}

func (s *KeywordContext) NOT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNOT, 0)
}

func (s *KeywordContext) NOTHING() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNOTHING, 0)
}

func (s *KeywordContext) NULL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNULL, 0)
}

func (s *KeywordContext) OBJECT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserOBJECT, 0)
}

func (s *KeywordContext) OPTIONAL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserOPTIONAL, 0)
}

func (s *KeywordContext) OR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserOR, 0)
}

func (s *KeywordContext) PARAMARRAY() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPARAMARRAY, 0)
}

func (s *KeywordContext) PRESERVE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPRESERVE, 0)
}

func (s *KeywordContext) PSET() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPSET, 0)
}

func (s *KeywordContext) PTRSAFE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPTRSAFE, 0)
}

func (s *KeywordContext) REM() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserREM, 0)
}

func (s *KeywordContext) SGN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSGN, 0)
}

func (s *KeywordContext) SINGLE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSINGLE, 0)
}

func (s *KeywordContext) SPC() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSPC, 0)
}

func (s *KeywordContext) STRING() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSTRING, 0)
}

func (s *KeywordContext) TAB() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserTAB, 0)
}

func (s *KeywordContext) TEXT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserTEXT, 0)
}

func (s *KeywordContext) THEN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserTHEN, 0)
}

func (s *KeywordContext) TO() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserTO, 0)
}

func (s *KeywordContext) TRUE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserTRUE, 0)
}

func (s *KeywordContext) TYPEOF() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserTYPEOF, 0)
}

func (s *KeywordContext) UBOUND() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserUBOUND, 0)
}

func (s *KeywordContext) UNTIL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserUNTIL, 0)
}

func (s *KeywordContext) VARIANT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserVARIANT, 0)
}

func (s *KeywordContext) VERSION() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserVERSION, 0)
}

func (s *KeywordContext) WITHEVENTS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserWITHEVENTS, 0)
}

func (s *KeywordContext) XOR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserXOR, 0)
}

func (s *KeywordContext) STEP() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSTEP, 0)
}

func (s *KeywordContext) ON_ERROR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserON_ERROR, 0)
}

func (s *KeywordContext) RESUME_NEXT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserRESUME_NEXT, 0)
}

func (s *KeywordContext) ERROR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserERROR, 0)
}

func (s *KeywordContext) APPEND() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserAPPEND, 0)
}

func (s *KeywordContext) BINARY() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserBINARY, 0)
}

func (s *KeywordContext) OUTPUT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserOUTPUT, 0)
}

func (s *KeywordContext) RANDOM() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserRANDOM, 0)
}

func (s *KeywordContext) ACCESS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserACCESS, 0)
}

func (s *KeywordContext) READ() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserREAD, 0)
}

func (s *KeywordContext) WRITE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserWRITE, 0)
}

func (s *KeywordContext) READ_WRITE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserREAD_WRITE, 0)
}

func (s *KeywordContext) SHARED() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSHARED, 0)
}

func (s *KeywordContext) LOCK_READ() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLOCK_READ, 0)
}

func (s *KeywordContext) LOCK_WRITE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLOCK_WRITE, 0)
}

func (s *KeywordContext) LOCK_READ_WRITE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLOCK_READ_WRITE, 0)
}

func (s *KeywordContext) LINE_INPUT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLINE_INPUT, 0)
}

func (s *KeywordContext) RESET() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserRESET, 0)
}

func (s *KeywordContext) WIDTH() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserWIDTH, 0)
}

func (s *KeywordContext) PRINT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPRINT, 0)
}

func (s *KeywordContext) GET() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserGET, 0)
}

func (s *KeywordContext) PUT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPUT, 0)
}

func (s *KeywordContext) CLOSE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCLOSE, 0)
}

func (s *KeywordContext) INPUT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserINPUT, 0)
}

func (s *KeywordContext) LOCK() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLOCK, 0)
}

func (s *KeywordContext) OPEN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserOPEN, 0)
}

func (s *KeywordContext) SEEK() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSEEK, 0)
}

func (s *KeywordContext) UNLOCK() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserUNLOCK, 0)
}

func (s *KeywordContext) NAME() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNAME, 0)
}

func (s *KeywordContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *KeywordContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) Keyword() (localctx IKeywordContext) {
	localctx = NewKeywordContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 56, VBAConditionalCompilationParserRULE_keyword)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(497)
		_la = p.GetTokenStream().LA(1)

		if !(((int64(_la) & ^0x3f) == 0 && ((int64(1) << _la) & 2278680789921036286) != 0) || ((int64((_la - 64)) & ^0x3f) == 0 && ((int64(1) << (_la - 64)) & 9116482636005507099) != 0) || ((int64((_la - 129)) & ^0x3f) == 0 && ((int64(1) << (_la - 129)) & -760485564683782209) != 0) || ((int64((_la - 193)) & ^0x3f) == 0 && ((int64(1) << (_la - 193)) & 4222124650674813) != 0)) {
			p.GetErrorHandler().RecoverInline(p)
		} else {
			p.GetErrorHandler().ReportMatch(p)
			p.Consume()
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IMarkerKeywordContext is an interface to support dynamic dispatch.
type IMarkerKeywordContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	AS() antlr.TerminalNode

	// IsMarkerKeywordContext differentiates from other interfaces.
	IsMarkerKeywordContext()
}

type MarkerKeywordContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyMarkerKeywordContext() *MarkerKeywordContext {
	var p = new(MarkerKeywordContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_markerKeyword
	return p
}

func InitEmptyMarkerKeywordContext(p *MarkerKeywordContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_markerKeyword
}

func (*MarkerKeywordContext) IsMarkerKeywordContext() {}

func NewMarkerKeywordContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *MarkerKeywordContext {
	var p = new(MarkerKeywordContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_markerKeyword

	return p
}

func (s *MarkerKeywordContext) GetParser() antlr.Parser { return s.parser }

func (s *MarkerKeywordContext) AS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserAS, 0)
}

func (s *MarkerKeywordContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *MarkerKeywordContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) MarkerKeyword() (localctx IMarkerKeywordContext) {
	localctx = NewMarkerKeywordContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 58, VBAConditionalCompilationParserRULE_markerKeyword)
	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(499)
		p.Match(VBAConditionalCompilationParserAS)
		if p.HasError() {
				// Recognition error - abort rule
				goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IStatementKeywordContext is an interface to support dynamic dispatch.
type IStatementKeywordContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	CALL() antlr.TerminalNode
	CASE() antlr.TerminalNode
	CONST() antlr.TerminalNode
	DECLARE() antlr.TerminalNode
	DEFBOOL() antlr.TerminalNode
	DEFBYTE() antlr.TerminalNode
	DEFCUR() antlr.TerminalNode
	DEFDATE() antlr.TerminalNode
	DEFDBL() antlr.TerminalNode
	DEFINT() antlr.TerminalNode
	DEFLNG() antlr.TerminalNode
	DEFLNGLNG() antlr.TerminalNode
	DEFLNGPTR() antlr.TerminalNode
	DEFOBJ() antlr.TerminalNode
	DEFSNG() antlr.TerminalNode
	DEFSTR() antlr.TerminalNode
	DEFVAR() antlr.TerminalNode
	DIM() antlr.TerminalNode
	DO() antlr.TerminalNode
	ELSE() antlr.TerminalNode
	ELSEIF() antlr.TerminalNode
	ENUM() antlr.TerminalNode
	ERASE() antlr.TerminalNode
	EVENT() antlr.TerminalNode
	EXIT() antlr.TerminalNode
	EXIT_DO() antlr.TerminalNode
	EXIT_FOR() antlr.TerminalNode
	EXIT_FUNCTION() antlr.TerminalNode
	EXIT_PROPERTY() antlr.TerminalNode
	EXIT_SUB() antlr.TerminalNode
	END_SELECT() antlr.TerminalNode
	END_WITH() antlr.TerminalNode
	FOR() antlr.TerminalNode
	FRIEND() antlr.TerminalNode
	FUNCTION() antlr.TerminalNode
	GLOBAL() antlr.TerminalNode
	GOSUB() antlr.TerminalNode
	GOTO() antlr.TerminalNode
	IF() antlr.TerminalNode
	IMPLEMENTS() antlr.TerminalNode
	LET() antlr.TerminalNode
	LOOP() antlr.TerminalNode
	LSET() antlr.TerminalNode
	NEXT() antlr.TerminalNode
	ON() antlr.TerminalNode
	OPTION() antlr.TerminalNode
	PRIVATE() antlr.TerminalNode
	PUBLIC() antlr.TerminalNode
	RAISEEVENT() antlr.TerminalNode
	REDIM() antlr.TerminalNode
	RESUME() antlr.TerminalNode
	RETURN() antlr.TerminalNode
	RSET() antlr.TerminalNode
	SELECT() antlr.TerminalNode
	SET() antlr.TerminalNode
	STATIC() antlr.TerminalNode
	STOP() antlr.TerminalNode
	SUB() antlr.TerminalNode
	TYPE() antlr.TerminalNode
	WEND() antlr.TerminalNode
	WHILE() antlr.TerminalNode
	WITH() antlr.TerminalNode

	// IsStatementKeywordContext differentiates from other interfaces.
	IsStatementKeywordContext()
}

type StatementKeywordContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyStatementKeywordContext() *StatementKeywordContext {
	var p = new(StatementKeywordContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_statementKeyword
	return p
}

func InitEmptyStatementKeywordContext(p *StatementKeywordContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_statementKeyword
}

func (*StatementKeywordContext) IsStatementKeywordContext() {}

func NewStatementKeywordContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *StatementKeywordContext {
	var p = new(StatementKeywordContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_statementKeyword

	return p
}

func (s *StatementKeywordContext) GetParser() antlr.Parser { return s.parser }

func (s *StatementKeywordContext) CALL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCALL, 0)
}

func (s *StatementKeywordContext) CASE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCASE, 0)
}

func (s *StatementKeywordContext) CONST() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserCONST, 0)
}

func (s *StatementKeywordContext) DECLARE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDECLARE, 0)
}

func (s *StatementKeywordContext) DEFBOOL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFBOOL, 0)
}

func (s *StatementKeywordContext) DEFBYTE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFBYTE, 0)
}

func (s *StatementKeywordContext) DEFCUR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFCUR, 0)
}

func (s *StatementKeywordContext) DEFDATE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFDATE, 0)
}

func (s *StatementKeywordContext) DEFDBL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFDBL, 0)
}

func (s *StatementKeywordContext) DEFINT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFINT, 0)
}

func (s *StatementKeywordContext) DEFLNG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFLNG, 0)
}

func (s *StatementKeywordContext) DEFLNGLNG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFLNGLNG, 0)
}

func (s *StatementKeywordContext) DEFLNGPTR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFLNGPTR, 0)
}

func (s *StatementKeywordContext) DEFOBJ() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFOBJ, 0)
}

func (s *StatementKeywordContext) DEFSNG() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFSNG, 0)
}

func (s *StatementKeywordContext) DEFSTR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFSTR, 0)
}

func (s *StatementKeywordContext) DEFVAR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDEFVAR, 0)
}

func (s *StatementKeywordContext) DIM() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDIM, 0)
}

func (s *StatementKeywordContext) DO() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserDO, 0)
}

func (s *StatementKeywordContext) ELSE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserELSE, 0)
}

func (s *StatementKeywordContext) ELSEIF() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserELSEIF, 0)
}

func (s *StatementKeywordContext) ENUM() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserENUM, 0)
}

func (s *StatementKeywordContext) ERASE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserERASE, 0)
}

func (s *StatementKeywordContext) EVENT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEVENT, 0)
}

func (s *StatementKeywordContext) EXIT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEXIT, 0)
}

func (s *StatementKeywordContext) EXIT_DO() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEXIT_DO, 0)
}

func (s *StatementKeywordContext) EXIT_FOR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEXIT_FOR, 0)
}

func (s *StatementKeywordContext) EXIT_FUNCTION() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEXIT_FUNCTION, 0)
}

func (s *StatementKeywordContext) EXIT_PROPERTY() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEXIT_PROPERTY, 0)
}

func (s *StatementKeywordContext) EXIT_SUB() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEXIT_SUB, 0)
}

func (s *StatementKeywordContext) END_SELECT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEND_SELECT, 0)
}

func (s *StatementKeywordContext) END_WITH() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserEND_WITH, 0)
}

func (s *StatementKeywordContext) FOR() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserFOR, 0)
}

func (s *StatementKeywordContext) FRIEND() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserFRIEND, 0)
}

func (s *StatementKeywordContext) FUNCTION() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserFUNCTION, 0)
}

func (s *StatementKeywordContext) GLOBAL() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserGLOBAL, 0)
}

func (s *StatementKeywordContext) GOSUB() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserGOSUB, 0)
}

func (s *StatementKeywordContext) GOTO() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserGOTO, 0)
}

func (s *StatementKeywordContext) IF() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserIF, 0)
}

func (s *StatementKeywordContext) IMPLEMENTS() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserIMPLEMENTS, 0)
}

func (s *StatementKeywordContext) LET() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLET, 0)
}

func (s *StatementKeywordContext) LOOP() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLOOP, 0)
}

func (s *StatementKeywordContext) LSET() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLSET, 0)
}

func (s *StatementKeywordContext) NEXT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserNEXT, 0)
}

func (s *StatementKeywordContext) ON() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserON, 0)
}

func (s *StatementKeywordContext) OPTION() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserOPTION, 0)
}

func (s *StatementKeywordContext) PRIVATE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPRIVATE, 0)
}

func (s *StatementKeywordContext) PUBLIC() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserPUBLIC, 0)
}

func (s *StatementKeywordContext) RAISEEVENT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserRAISEEVENT, 0)
}

func (s *StatementKeywordContext) REDIM() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserREDIM, 0)
}

func (s *StatementKeywordContext) RESUME() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserRESUME, 0)
}

func (s *StatementKeywordContext) RETURN() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserRETURN, 0)
}

func (s *StatementKeywordContext) RSET() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserRSET, 0)
}

func (s *StatementKeywordContext) SELECT() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSELECT, 0)
}

func (s *StatementKeywordContext) SET() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSET, 0)
}

func (s *StatementKeywordContext) STATIC() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSTATIC, 0)
}

func (s *StatementKeywordContext) STOP() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSTOP, 0)
}

func (s *StatementKeywordContext) SUB() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserSUB, 0)
}

func (s *StatementKeywordContext) TYPE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserTYPE, 0)
}

func (s *StatementKeywordContext) WEND() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserWEND, 0)
}

func (s *StatementKeywordContext) WHILE() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserWHILE, 0)
}

func (s *StatementKeywordContext) WITH() antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserWITH, 0)
}

func (s *StatementKeywordContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *StatementKeywordContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) StatementKeyword() (localctx IStatementKeywordContext) {
	localctx = NewStatementKeywordContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 60, VBAConditionalCompilationParserRULE_statementKeyword)
	var _la int

	p.EnterOuterAlt(localctx, 1)
	{
		p.SetState(501)
		_la = p.GetTokenStream().LA(1)

		if !(((int64((_la - 22)) & ^0x3f) == 0 && ((int64(1) << (_la - 22)) & 9223250540819907585) != 0) || ((int64((_la - 87)) & ^0x3f) == 0 && ((int64(1) << (_la - 87)) & 585752737811966211) != 0) || ((int64((_la - 160)) & ^0x3f) == 0 && ((int64(1) << (_la - 160)) & 12112161903137) != 0)) {
			p.GetErrorHandler().RecoverInline(p)
		} else {
			p.GetErrorHandler().ReportMatch(p)
			p.Consume()
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


// IWhiteSpaceContext is an interface to support dynamic dispatch.
type IWhiteSpaceContext interface {
	antlr.ParserRuleContext

	// GetParser returns the parser.
	GetParser() antlr.Parser

	// Getter signatures
	AllWS() []antlr.TerminalNode
	WS(i int) antlr.TerminalNode
	AllLINE_CONTINUATION() []antlr.TerminalNode
	LINE_CONTINUATION(i int) antlr.TerminalNode

	// IsWhiteSpaceContext differentiates from other interfaces.
	IsWhiteSpaceContext()
}

type WhiteSpaceContext struct {
	antlr.BaseParserRuleContext
	parser antlr.Parser
}

func NewEmptyWhiteSpaceContext() *WhiteSpaceContext {
	var p = new(WhiteSpaceContext)
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_whiteSpace
	return p
}

func InitEmptyWhiteSpaceContext(p *WhiteSpaceContext)  {
	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, nil, -1)
	p.RuleIndex = VBAConditionalCompilationParserRULE_whiteSpace
}

func (*WhiteSpaceContext) IsWhiteSpaceContext() {}

func NewWhiteSpaceContext(parser antlr.Parser, parent antlr.ParserRuleContext, invokingState int) *WhiteSpaceContext {
	var p = new(WhiteSpaceContext)

	antlr.InitBaseParserRuleContext(&p.BaseParserRuleContext, parent, invokingState)

	p.parser = parser
	p.RuleIndex = VBAConditionalCompilationParserRULE_whiteSpace

	return p
}

func (s *WhiteSpaceContext) GetParser() antlr.Parser { return s.parser }

func (s *WhiteSpaceContext) AllWS() []antlr.TerminalNode {
	return s.GetTokens(VBAConditionalCompilationParserWS)
}

func (s *WhiteSpaceContext) WS(i int) antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserWS, i)
}

func (s *WhiteSpaceContext) AllLINE_CONTINUATION() []antlr.TerminalNode {
	return s.GetTokens(VBAConditionalCompilationParserLINE_CONTINUATION)
}

func (s *WhiteSpaceContext) LINE_CONTINUATION(i int) antlr.TerminalNode {
	return s.GetToken(VBAConditionalCompilationParserLINE_CONTINUATION, i)
}

func (s *WhiteSpaceContext) GetRuleContext() antlr.RuleContext {
	return s
}

func (s *WhiteSpaceContext) ToStringTree(ruleNames []string, recog antlr.Recognizer) string {
	return antlr.TreesStringTree(s, ruleNames, recog)
}




func (p *VBAConditionalCompilationParser) WhiteSpace() (localctx IWhiteSpaceContext) {
	localctx = NewWhiteSpaceContext(p, p.GetParserRuleContext(), p.GetState())
	p.EnterRule(localctx, 62, VBAConditionalCompilationParserRULE_whiteSpace)
	var _la int

	var _alt int

	p.EnterOuterAlt(localctx, 1)
	p.SetState(504)
	p.GetErrorHandler().Sync(p)
	if p.HasError() {
		goto errorExit
	}
	_alt = 1
	for ok := true; ok; ok = _alt != 2 && _alt != antlr.ATNInvalidAltNumber {
		switch _alt {
		case 1:
				{
					p.SetState(503)
					_la = p.GetTokenStream().LA(1)

					if !(_la == VBAConditionalCompilationParserWS || _la == VBAConditionalCompilationParserLINE_CONTINUATION) {
						p.GetErrorHandler().RecoverInline(p)
					} else {
						p.GetErrorHandler().ReportMatch(p)
						p.Consume()
					}
				}




		default:
			p.SetError(antlr.NewNoViableAltException(p, nil, nil, nil, nil, nil))
			goto errorExit
		}

		p.SetState(506)
		p.GetErrorHandler().Sync(p)
		_alt = p.GetInterpreter().AdaptivePredict(p.BaseParser, p.GetTokenStream(), 58, p.GetParserRuleContext())
		if p.HasError() {
			goto errorExit
		}
	}



errorExit:
	if p.HasError() {
		v := p.GetError()
		localctx.SetException(v)
		p.GetErrorHandler().ReportError(p, v)
		p.GetErrorHandler().Recover(p, v)
		p.SetError(nil)
	}
	p.ExitRule()
	return localctx
	goto errorExit // Trick to prevent compiler error if the label is not used
}


func (p *VBAConditionalCompilationParser) Sempred(localctx antlr.RuleContext, ruleIndex, predIndex int) bool {
	switch ruleIndex {
	case 18:
			var t *CcExpressionContext = nil
			if localctx != nil { t = localctx.(*CcExpressionContext) }
			return p.CcExpression_Sempred(t, predIndex)


	default:
		panic("No predicate with index: " + fmt.Sprint(ruleIndex))
	}
}

func (p *VBAConditionalCompilationParser) CcExpression_Sempred(localctx antlr.RuleContext, predIndex int) bool {
	switch predIndex {
	case 0:
			return p.Precpred(p.GetParserRuleContext(), 17)

	case 1:
			return p.Precpred(p.GetParserRuleContext(), 15)

	case 2:
			return p.Precpred(p.GetParserRuleContext(), 14)

	case 3:
			return p.Precpred(p.GetParserRuleContext(), 13)

	case 4:
			return p.Precpred(p.GetParserRuleContext(), 12)

	case 5:
			return p.Precpred(p.GetParserRuleContext(), 11)

	case 6:
			return p.Precpred(p.GetParserRuleContext(), 10)

	case 7:
			return p.Precpred(p.GetParserRuleContext(), 8)

	case 8:
			return p.Precpred(p.GetParserRuleContext(), 7)

	case 9:
			return p.Precpred(p.GetParserRuleContext(), 6)

	case 10:
			return p.Precpred(p.GetParserRuleContext(), 5)

	case 11:
			return p.Precpred(p.GetParserRuleContext(), 4)

	default:
		panic("No predicate with index: " + fmt.Sprint(predIndex))
	}
}

