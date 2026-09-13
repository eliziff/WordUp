param([Parameter(Mandatory=$true)][string]$AntlrJar)
$ErrorActionPreference = 'Stop'
[Diagnostics.Process]::GetCurrentProcess().PriorityClass = 'BelowNormal'
$destination = Join-Path $PSScriptRoot 'generated'
$utf8 = New-Object Text.UTF8Encoding($false)
$parser = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'grammar/VBAParser.g4'))
$parser = $parser.Replace('superClass = VBABaseParser;', '').Replace('contextSuperClass = VBABaseParserRuleContext;', '')
$parser = [regex]::Replace($parser, '\{([^{}]+)\}\?', {
    param($match)
    $code = [regex]::Replace($match.Groups[1].Value, '\b(MatchesRegex|TextOf|TokenAtRelativePosition|EqualsString|EqualsStringIgnoringCase|IsTokenType|TokenTypeAtRelativePosition)\(', 'p.$1(')
    $code = [regex]::Replace($code, '\b(DOEVENTS|END|CLOSE|ELSE|LOOP|NEXT|RANDOMIZE|REM|RESUME|RETURN|STOP|WEND|IDENTIFIER|L_SQUARE_BRACKET|NEWLINE|LINE_CONTINUATION|LPAREN)\b', 'VBAParser$1')
    '{' + $code + '}?'
})
# ANTLR 4.13 rejects EOF inside this closure. Parse appends a final newline.
$parser = $parser.Replace('individualNonEOFEndOfStatement+ | whiteSpace? EOF', 'individualNonEOFEndOfStatement+').Replace('(NEWLINE | EOF)', 'NEWLINE')
$lexer = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'grammar/VBALexer.g4'))
$lexer = $lexer.Replace('superClass = VBABaseLexer;', '').Replace('contextSuperClass = VBABaseParser;', '').Replace('IsChar(', 'p.IsChar(').Replace('CharAtRelativePosition(', 'p.CharAtRelativePosition(')
[IO.File]::WriteAllText((Join-Path $destination 'VBAParser.g4'), $parser, $utf8)
[IO.File]::WriteAllText((Join-Path $destination 'VBALexer.g4'), $lexer, $utf8)
$jar = (Resolve-Path -LiteralPath $AntlrJar).Path
Push-Location $destination
try {
    & java -Xmx256m -jar $jar -encoding UTF-8 '-Dlanguage=Go' -package generated -no-listener -Xexact-output-dir -o $destination VBALexer.g4 VBAParser.g4
    if ($LASTEXITCODE -ne 0) { throw 'ANTLR generation failed' }
    if (!(Test-Path 'vba_parser.go')) { throw 'ANTLR did not produce vba_parser.go' }
    $conditional=Join-Path $PSScriptRoot 'conditional'
    New-Item -ItemType Directory -Force $conditional | Out-Null
    $grammar=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'grammar/VBAConditionalCompilationParser.g4')).Replace('(NEWLINE | EOF)','NEWLINE')
    [IO.File]::WriteAllText((Join-Path $destination 'VBAConditionalCompilationParser.g4'),$grammar,$utf8)
    & java -Xmx256m -jar $jar -encoding UTF-8 '-Dlanguage=Go' -package conditional -no-listener -Xexact-output-dir -o $conditional VBAConditionalCompilationParser.g4
    if($LASTEXITCODE -ne 0){throw 'Conditional parser generation failed'}
} finally { Pop-Location }
