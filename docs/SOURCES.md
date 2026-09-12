# Implementation references

Primary sources consulted; the code does not establish interoperability merely by citing these.

- Microsoft MS-CFB: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
- Microsoft MS-OVBA: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
- Microsoft MS-OFORMS: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/
- CreateDesktopW: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-createdesktopw
- STARTUPINFOW: https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/ns-processthreadsapi-startupinfow
- AccessibleObjectFromWindow (Word _WwG native Window object): https://learn.microsoft.com/en-us/windows/win32/api/oleacc/nf-oleacc-accessibleobjectfromwindow
- Word AutomationSecurity: https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity
- Word Page.EnhMetaFileBits: https://learn.microsoft.com/en-us/office/vba/api/word.page.enhmetafilebits
- VBA compiler constants: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/compiler-constants
- Apple Event Manager: https://developer.apple.com/documentation/applicationservices/apple_event_manager
- MCP 2025-11-25: https://modelcontextprotocol.io/specification/2025-11-25
- pyOpenVBA forms persistence reference: https://github.com/WilliamSmithEdward/pyOpenVBA (MIT; see THIRD-PARTY-NOTICES.md)

No external service is called by the harness to implement source editing or native host control. Native macros themselves may call external services if the authorized user/agent programs them to do so.
