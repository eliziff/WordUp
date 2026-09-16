package component

import _ "embed"

// Bundled component VBA lives in ordinary .bas files under vba/ so it can be
// read, diffed, and loaded into Word exactly as shipped; Go only embeds the
// bytes. Keep the files LF-terminated (see .gitattributes): the embedded text
// is hashed into component manifests and installed workspace locks.

//go:embed vba/WordUpSafeEdit.bas
var safeEditSource string

//go:embed vba/WordUpFormShell.bas
var formShellSource string

//go:embed vba/WordUpProgress.bas
var progressSource string

//go:embed vba/WordUpRibbon.bas
var ribbonSource string

//go:embed vba/WordUpHotkey.bas
var hotkeySource string

//go:embed vba/WordUpContextMenu.bas
var contextMenuSource string

//go:embed vba/WordUpStyleConverter.bas
var styleConverterSource string

//go:embed vba/WordUpFieldRefresh.bas
var fieldRefreshSource string

//go:embed vba/WordUpTextOperations.bas
var textOperationsSource string
