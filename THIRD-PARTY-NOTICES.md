# Third-party notices

## pyOpenVBA design/format reference

The native MS-OFORMS implementation was developed with reference to Microsoft specifications and the pyOpenVBA project's record tables, persisted control metadata and MultiPage bookkeeping. No Python package or pyOpenVBA runtime is required or bundled. Adapted format tables/constants and substantial corresponding implementation ideas retain this MIT notice.

Source: https://github.com/WilliamSmithEdward/pyOpenVBA
Consulted forms.py SHA 2c7994bc31facf7c1c768cf3f088cdc7bc2f4ca2;
_oforms_pages.py SHA aeba166c177afcf8191a0eeffb9e3e246d52a054.

MIT License

Copyright (c) 2026 William Smith

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Go runtime and standard library

Native executables statically include the Go runtime/standard library. This is not a zero-runtime C program; there is no separately installed Go runtime.

Copyright (c) 2009 The Go Authors. All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

   * Redistributions of source code must retain the above copyright
notice, this list of conditions and the following disclaimer.
   * Redistributions in binary form must reproduce the above
copyright notice, this list of conditions and the following disclaimer
in the documentation and/or other materials provided with the
distribution.
   * Neither the name of Google Inc. nor the names of its
contributors may be used to endorse or promote products derived from
this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

## Microsoft Office and operating systems

Microsoft Word, Office components, system DLLs/frameworks, fonts, and third-party ActiveX controls are not redistributed. The app calls existing installed system and Office APIs. Microsoft/Apple trademarks belong to their respective owners. The user's private ALR template is not in the generic source distribution.


## Microsoft RibbonX schemas

The embedded 2007 and 2010 schemas originate from OfficeDev/office-custom-ui-editor, obtained through Office RibbonX Editor.


License notice for OfficeDev/office-custom-ui-editor
----------------------------------------------------

MIT License

Copyright (c) Microsoft Corporation. All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE

## Rubberduck grammar and Go parser dependencies

The integrated VBA grammar is adapted from [Rubberduck](https://github.com/rubberduck-vba/Rubberduck), observed commit `fae50adab188126a5e7d2a1cefc3328cc18af482`. Original grammar and predicate inputs, attribution, and GPL-3.0-or-later terms are retained in `internal/vbaparse/grammar`; regeneration is documented in `internal/vbaparse/README.md`. The combined distribution uses the root GPL license; the original MIT notice is retained in LICENSE-MIT-ORIGINAL.

ANTLR Go runtime 4.13.1 retains its BSD notice in `internal/vbaparse/ANTLR-LICENSE`. The pinned golang.org/x/exp dependency retains its BSD notice in `internal/vbaparse/X-EXP-LICENSE`. Exact Go versions/checksums are in go.mod/go.sum.

## XPath evidence queries

`xml.query` uses antchfx/xmlquery 1.5.1 and antchfx/xpath 1.3.6 (MIT).
Their exact upstream notices and the transitive groupcache, x/net and x/text
license/patent texts are retained in `internal/office/XPATH-NOTICES` and included
by the Windows packaging script. Dependency checksums are in go.sum.

## RibbonX callback metadata

Adapted Office RibbonX Editor material retains Fernando Andreu's MIT notice in `internal/office/RIBBONX-LICENSE`, in addition to the Microsoft schema notice above.

## Optional OfficeTools helper

The helper uses Microsoft Open XML SDK and Framework 3.3.0 (MIT), FlaUI Core/UIA3 4.0.0 (MIT), Interop.UIAutomationClient 10.19041.0, and System.Management 5.0.0. Exact targets are listed in tools/office-bridge/build.ps1. Package nuspec metadata and package-provided license/notice files are copied beside the bundled DLLs. Open XML SDK's upstream v3.3.0 license is vendored at `tools/office-bridge/OPENXML-LICENSE` and copied into the helper bundle because the NuGet packages supply a license expression rather than the full text. Source: https://github.com/dotnet/Open-XML-SDK/blob/v3.3.0/LICENSE.

This inventory does not establish that all optional binary bundles are release-ready. In particular, the oletools/PyInstaller transitive bundle, its licenses, reproducible dependency pins, and corresponding source delivery still require a release audit.

### Legal Structure Parser marker adaptation

`internal/structure/WordUpStructure.bas` adapts journal label recognition from legal-structure `src/journal.rs`, revision f73bcd5d66575c0504e97fc2c3b95bf5589377e5. Copyright (c) 2026 legal-structure contributors. MIT license: `internal/structure/LICENSE`. The remaining Word adapter code collects native evidence; it does not include the full upstream inference engine.

`internal/structure/ladder.go` and marker interpretation adapt Legal PDF Parser support primitives at e1b060bdf9b92b9f6295bfc1e7933bd5a446af22. See `internal/structure/PDF-LICENSE` and `UPSTREAM.md` for license and adaptation scope.
