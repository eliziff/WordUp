# Reusing legal-document primitives

WordUp points agents to two complementary repositories:

- [Legal Structure Parser](https://github.com/eliziff/legal-structure-parser): provider-neutral structure, native versus inferred evidence, numbering candidates, journal headings, footnotes, exact queries and Unicode offset conversion.
- [Legal PDF Parser](https://github.com/eliziff/legal-pdf-parser): digital PDF extraction, geometry, reading order, superscript evidence, page furniture and source witnesses. Enable OCR only when a scanned source requires it.

These are optional agent-side tools. WordUp does not clone, download, build or load them automatically. The initial integration is documentation plus links in generated workspace instructions and the help response. No Rust dependency, FFI layer, bundled OCR model or additional end-user setup is introduced.

Use native Word styles, outline levels, numbering, stories and ranges as authoritative inputs. Treat visual or textual classification as inference and retain its evidence. Preserve the difference between UTF-16 Word positions and scalar/byte offsets in other engines. Do not turn numbered quotations, footnote markers or page headers into inferred document headings.

For each adaptation, record the upstream revision, exact source primitive, applicable license and tests. Keep the format-specific Word adapter small. Prefer using an existing executable or binding during authoring when available; port a bounded rule into VBA only when it needs to travel inside the standalone DOTM. Preserve upstream notices for copied code.

If repeated projects justify deeper integration, an optional adapter exchanging bounded, source-hashed structure records is the next step. That boundary should accept authoritative Word evidence and return candidates with provenance; it should not silently rewrite documents. A maintained native library binding is justified only after that reusable boundary is demonstrated.
