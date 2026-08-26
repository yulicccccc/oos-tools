import ast
import glob
import sys
import sysconfig

sys.stdout.reconfigure(encoding='utf-8')

py_files = glob.glob("*.py") + glob.glob("pages/*.py")

imports = set()
for fname in py_files:
    try:
        with open(fname, "r", encoding="utf-8") as f:
            node = ast.parse(f.read(), filename=fname)
        for n in ast.walk(node):
            if isinstance(n, ast.Import):
                for alias in n.names:
                    imports.add(alias.name.split('.')[0])
            elif isinstance(n, ast.ImportFrom):
                if n.module:
                    imports.add(n.module.split('.')[0])
    except Exception as e:
        print(f"Error parsing {fname}: {e}")

local_modules = {f.replace('.py', '') for f in glob.glob("*.py")}
stdlib = set(sys.builtin_module_names) | {
    'os', 'sys', 're', 'json', 'io', 'subprocess', 'time', 'datetime', 'math', 
    'pathlib', 'shutil', 'glob', 'argparse', 'ast', 'py_compile', 'codecs', 'copy',
    'collections', 'typing', 'functools', 'itertools', 'random', 'logging', 'tempfile'
}

third_party = imports - stdlib - local_modules
print("All Third-Party Imports Detected Across Codebase:")
for imp in sorted(third_party):
    print(f"  - {imp}")
