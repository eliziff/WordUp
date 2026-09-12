# WordUp: the product contract

WordUp should let any coding agent author, edit, test and deliver native Word
templates as fluently as it edits text files. The user downloads one app, points
an agent at it, and does not import modules, debug VBA, configure a VM or register
for a service. Installed Microsoft Word supplies the actual runtime.

The scope includes arbitrary VBA, persistent UserForms, RibbonX, styles,
numbering, document composition and reusable building blocks. A DOCX or image
can guide the agent's design. Results must remain editable native Word content.
The agent must observe behavior and layout, iterate quickly, preserve unrelated
content, and install a tested result for the next Word launch. Mac portability
and actual Mac testing are a stretch goal.

## Acceptance gates

| Gate | Required proof |
|---|---|
| Source editing | Import/build identity; add, edit and delete modules; preserve unrelated parts and unsupported binary data. |
| Native execution | Run generated VBA in an owned Word instance; assert results and deliberate counterexamples. |
| Forms and Ribbon | Instantiate saved nested forms, operate their controls, observe events and real Ribbon callbacks. |
| Document design | Assert styles, content controls, footnotes and saved parts; inspect actual Word page renders. |
| Isolation | Keep the user's existing Word process/documents intact; contain timeouts and modal failures; leave no owned Word process. |
| Fast iteration | Measure cold startup and repeated warm builds/calls on actual Word, with workload and host recorded. |
| Delivery | Bind acceptance to exact artifact bytes, back up the target, reject stale writes, and recover pending activation after login. |
| Mac stretch | Execute on Word for Mac; Windows passes and static portability checks cannot substitute. |

## Starting gaps observed in the supplied project

The supplied source has a substantial native implementation, but its historical
evidence establishes offline Linux behavior only. The Windows/Mac binaries were
cross-compiled. Native startup, VBA compilation, forms, Ribbon, rendering,
coexistence and cleanup had never been exercised by the supplied validation.

Concrete missing capabilities include automatic reboot recovery for deployment,
Mac UI/render/compile parity, universal custom ActiveX serialization, and native
latency measurements. Image-based design depends on the external agent's vision;
the executable does not contain a model. The existing integrated self-test does
not require a whole-project compiler observation or a rendered page, so even a
pass would not establish the whole contract above.

These are engineering gaps to close, not permission to weaken the product.
Historical evidence in `evidence/` and `VALIDATION.md` is retained unchanged and
does not certify the renamed or subsequently modified executable.

## Rename compatibility

The product and executable are WordUp / `wordup`. The Go module is
`github.com/eliziff/WordUp`. Existing format-2 workspaces retain the `.wordwright`
storage directory so the rename does not strand imported templates. Stop an old
local session before opening its workspace with the new executable.

## Current status (0.3.0)

The starting gaps above describe the imported preview. Windows native compilation, forms, Ribbon, rendering, containment and latency now have real evidence. Login recovery and guarded restore have native delivery evidence. See VALIDATION-WINDOWS.md and LIMITS.md for current coverage and remaining gaps. Optional human pointing/feedback, improved ALR headings and Mac parity remain future work.
