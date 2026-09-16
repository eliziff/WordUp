# Benchmark 20260916-082631-baseline

model gpt-5.6-luna effort xhigh; exe build\bench-baseline\wordup.exe; timeout 1800s per run

| task | arm | runs | passed | median wall s | median commands | median input tok | median output tok | median suite ms |
|---|---|---:|---:|---:|---:|---:|---:|---:|
| alr-hotkey-unbind | components | 2 | 2 | 650.9 | 41.5 | 3958981.0 | 26587.5 | 4263.5 |
| alr-hotkey-unbind | core | 2 | 2 | 560.6 | 45.0 | 3404731.0 | 24554.0 | 4358.3 |
| alr-screenupdating-leak | components | 2 | 2 | 543.5 | 42.5 | 2926627.5 | 24510.5 | 4192.3 |
| alr-screenupdating-leak | core | 2 | 2 | 668.9 | 56.5 | 3582753.5 | 27599.5 | 4183.4 |
| studio-italicize-supra-notes | components | 2 | 0 | 785.5 | 44.0 | 4613376.5 | 34793.0 | 2834.6 |
| studio-italicize-supra-notes | core | 2 | 0 | 847.2 | 37.0 | 3150682.5 | 33620.5 | 2769.2 |

## Runs

- alr-hotkey-unbind / core / run-1: PASS; wall 457.4 s; commands 31; timed_out False; components []
- alr-hotkey-unbind / core / run-2: PASS; wall 663.9 s; commands 59; timed_out False; components []
- alr-hotkey-unbind / components / run-1: PASS; wall 513.6 s; commands 33; timed_out False; components []
- alr-hotkey-unbind / components / run-2: PASS; wall 788.2 s; commands 50; timed_out False; components []
- alr-screenupdating-leak / core / run-1: PASS; wall 568.1 s; commands 62; timed_out False; components []
- alr-screenupdating-leak / core / run-2: PASS; wall 769.6 s; commands 51; timed_out False; components []
- alr-screenupdating-leak / components / run-1: PASS; wall 523.5 s; commands 48; timed_out False; components []
- alr-screenupdating-leak / components / run-2: PASS; wall 563.4 s; commands 37; timed_out False; components []
- studio-italicize-supra-notes / core / run-1: FAIL package/**/customUI*.xml contains 'ItalicizeSupraInNotes'; wall 887.0 s; commands 43; timed_out False; components []
- studio-italicize-supra-notes / core / run-2: FAIL package/**/customUI*.xml contains 'ItalicizeSupraInNotes'; wall 807.5 s; commands 31; timed_out False; components []
- studio-italicize-supra-notes / components / run-1: FAIL package/**/customUI*.xml contains 'ItalicizeSupraInNotes'; wall 798.5 s; commands 48; timed_out False; components []
- studio-italicize-supra-notes / components / run-2: FAIL package/**/customUI*.xml contains 'ItalicizeSupraInNotes'; wall 772.5 s; commands 40; timed_out False; components []
