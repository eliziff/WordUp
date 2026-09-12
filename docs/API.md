# Agent API

All commands are local. Use `wordwright -w WORKSPACE call METHOD @parameters.json` for one-shot operations and `rpc METHOD @parameters.json` for a previously started warm session. `serve` accepts newline-delimited `{id,method,params}` JSON. `mcp` is stdio JSON-RPC with initialization and tools/call. Stdout contains protocol data; errors are structured and command failure is nonzero.

## Direct native object access (Windows)

Create `operation.json`:

```json
{"operation":{"op":"new","as":"doc"}}
```

Then use `rpc native.call @operation.json`. Subsequent operations retain object handles:

```json
{"operation":{"op":"get","target":"doc","member":"Content","as":"body"}}
```

```json
{"operation":{"op":"put","target":"body","member":"Text","value":"An actual Word document.\r"}}
```

```json
{"operation":{"op":"get","target":"body","member":"Text"}}
```

Methods accept positional `args` and named `named` parameters. Pass a retained object using `{"object":"body"}` and an omitted optional COM parameter using `{"missing":true}`. Object-returning calls require an `as` handle. Default application handle is `app`. Prefer document/Range objects to changing the user's application-wide defaults. `batch` contains an ordered `steps` array and avoids a process launch for every member access. Before opening a changed revision of a staged filename, close the prior document with `unload`, passing both its `target` handle and its original `file` path; this explicitly releases the staged-version guard. Keep filenames unique for simultaneously loaded revisions.

## Scratch native VBA

```json
{"operation":{"op":"eval","value":"Dim d As Document\nSet d = Documents.Add\nd.Content.Text = \"Native VBA wrote this.\"\nEvaluate = d.Paragraphs.Count\nd.Close SaveChanges:=wdDoNotSaveChanges"}}
```

This generates a separate temporary DOTM containing a real `Public Function Evaluate() As Variant`, runs it through Word, and returns the result plus generated source/hash. It does not modify the target template. Use a normal source component for persistent behavior. Native errors remain errors; the harness does not evaluate the body itself.

## Modal forms and UI

```json
{"operation":{"op":"begin","steps":[{"op":"run","macro":"ProjectName.ModuleName.ShowModalForm"}]}}
```

Use `poll` with `value` equal to the returned task token. While the macro is blocked, `ui.windows` and `ui.tree` can inspect owned windows, `ui.invoke` can invoke an accessible control, and `ui.set_value` can fill a field. Use actual HWND and accessibility child identifiers from that inspection. No guessed controls or global keyboard input. `ui.capture` captures an owned window. When finished, use `forget` to release a completed task record. A process has a bounded task budget.

## Native render and compile

```json
{"operation":{"op":"render","target":"doc","file":"C:\\Work\\MyTemplate\\reports\\pages","value":144}}
```

`child` optionally selects a 1-based page. Native PNG rendering is Windows-only. `format:"pdf"` requests Word's own PDF exporter. For actual compile observation, pass `op:"compile"`, a document `target`, and `member` with the project name for menu-target verification if VBProject access is unavailable. Use `begin` for a call that can block on compiler diagnostics and inspect the owned VBE accessibility tree. A disabled command is not a witnessed compiler pass.

## Acceptance suites

```json
{
  "schema":1,
  "name":"Citation behavior",
  "platforms":["windows"],
  "steps":[
    {"name":"Load exact candidate","operation":{"op":"open","file":"$artifact","as":"candidate"},"assert":[{"kind":"equals","path":"/open_and_repair","expected":false}]},
    {"name":"Execute native acceptance","operation":{"op":"run","macro":"$project.Tests.CheckCitation"},"assert":[{"kind":"equals","expected":"CITATION_OK"}]}
  ]
}
```

`$artifact`, `$project`, `$filename` are resolved from the exact staged candidate. Tests must include assertions. `path` uses an RFC 6901 JSON pointer into the actual returned value; empty path means the entire result. Kinds: `equals`, `not_equals`, `contains`, `matches`, `near` with tolerance, and `greater_than`. A suite can require a separately observed compile via `require_compile:true`. A raw native call can execute arbitrary code; an acceptance suite should state observable outcomes rather than trusting the generated code's claims alone.

The CLI `test` command uses a fresh process. Via agent `test`, set `fresh:true` for release acceptance; warm runs are useful while developing but cannot authorize deployment.

## Deployment

`deployment.json`:

```json
{"path":"dist/MyTemplate.dotm","reference":"reports/acceptance.json","output":"C:\\Templates\\MyTemplate.dotm"}
```

Use authorized `call deploy @deployment.json`. The app checks the exact artifact/suite/observations, stages the bytes, backs up the old file and either activates or waits locally for Word to exit. An interrupted plan remains at the returned `plan` path. `call activate` with `path` equal to that plan retries it; no OS startup persistence is installed.

## Images and style references

`reference.document` accepts a DOCX/DOTM path and returns native style definitions and paragraph/direct-formatting observations. It is not a complete resolved cascading-style engine. `reference.image` supplies actual image pixels to MCP/vision clients. An ordinary CLI agent can also read that same local image file. `image.compare` takes `reference`, `path`, optional `tolerance` and optional `output` for a difference PNG. Different dimensions fail: there is no hidden alignment or rescaling that could create a misleading pass.

## Mac native prototype

Use `dictionary` to inspect the installed Word SDEF and the implemented operation surface. Native `get`, `put`, `invoke`, `run` and `ae.send` use actual Apple Events. Event/property codes are discovered from the installed dictionary, not guessed from Windows names. Mac UI/render/full-compile operations are unsupported and must not be silently rerouted to Windows or emulated. `compat` is advisory source review; only actual Mac Word evidence can verify Mac behavior.
