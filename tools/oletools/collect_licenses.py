"""Collect installed build-environment provenance without network access."""
import hashlib
import importlib.metadata as metadata
import json
from pathlib import Path
import re
import shutil
import sys

output = Path(sys.argv[1])
output.mkdir(parents=True, exist_ok=True)
records = []
for dist in sorted(metadata.distributions(), key=lambda d: d.metadata['Name'].lower()):
    name = dist.metadata['Name']
    folder = output / re.sub(r'[^a-zA-Z0-9_.-]', '_', name)
    folder.mkdir(exist_ok=True)
    (folder / 'METADATA.txt').write_text(dist.read_text('METADATA') or '', encoding='utf-8')
    files = []
    for entry in dist.files or []:
        basename = Path(str(entry)).name
        if not basename.lower().startswith(('license', 'copying', 'notice')):
            continue
        source = Path(dist.locate_file(entry))
        if not source.is_file():
            continue
        target = folder / (str(len(files)) + '-' + basename)
        shutil.copyfile(source, target)
        files.append({'file': str(target.relative_to(output)), 'sha256': hashlib.sha256(target.read_bytes()).hexdigest()})
    if not files:
        vendored = Path(__file__).parent / 'licenses' / (name + '-LICENSE')
        if vendored.is_file():
            target = folder / 'LICENSE'
            shutil.copyfile(vendored, target)
            files.append({'file': str(target.relative_to(output)), 'sha256': hashlib.sha256(target.read_bytes()).hexdigest(), 'source': 'vendored upstream license'})
    records.append({'name': name, 'version': dist.version, 'license_files': files,
                    'license_files_missing': not bool(files)})
python_license = Path(sys.base_prefix) / 'LICENSE.txt'
if not python_license.is_file():
    raise SystemExit('Python runtime LICENSE.txt missing: ' + str(python_license))
shutil.copyfile(python_license, output / 'PYTHON-LICENSE.txt')
(output / 'inventory.json').write_text(json.dumps({'python': sys.version, 'scope': 'installed build environment, not an exact frozen-import inventory', 'distributions': records}, indent=2), encoding='utf-8')
print(json.dumps({'distributions': len(records), 'missing_license_files': [r['name'] for r in records if r['license_files_missing']]}))
