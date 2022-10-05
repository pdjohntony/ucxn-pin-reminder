import os
import sys
import importlib
from rich import print
from datetime import date
import time
import subprocess
import _version

og_exe_name   = "ucxn-pin-reminder" # name only, no .exe
spec_file     = "pyinstaller.spec"
dist_folder   = "dist"
venv_activate = "env\\Scripts\\activate.bat"
today         = date.today().strftime("%m-%d-%Y")
cwd           = os.path.dirname(os.path.abspath(__name__))
venv_activate = os.path.join(cwd, venv_activate)
dist_folder   = os.path.join(cwd, dist_folder)

def version_writer(n_version = _version.__version__):
	with open("_version.py", "w") as version_file:
		n_build = int(_version.__build__) + 1 # build always increases
		n_build = str(n_build).zfill(4) # pad with leading zeroes
		version_file.write(f"__version__    = '{n_version}'\n")
		version_file.write(f"__build__      = '{n_build}'\n")
		version_file.write(f"__build_date__ = '{today}'")

def version_printer(title = "--Current Build Info--"):
	importlib.reload(_version) # Reload the file
	print("\n"+("#" * 22))
	print(f"{title}")
	print(f"Version: {_version.__version__:>13}")
	print(f"Build: {_version.__build__:>15}")
	print(f"Build Date: {_version.__build_date__}")
	print("#" * 22)

#* 1 - Version
print(f"Activating py venv at '{venv_activate}'")
subprocess.run(venv_activate)

version_printer()

uinput = None
n_version = _version.__version__
ma, mi, p = _version.__version__.split(".")
while True:
	if uinput == None:
		uinput = input("\nIs this a major (ma), minor (mi), patch (p), or build only (b) (default) release?\n> ").lower() or "b"
	elif uinput == "ma":
		print(f"Incrementing major (ma) version")
		n_version = f"{int(ma)+1}.0.0"
		break
	elif uinput == "mi":
		print(f"Incrementing minor (mi) version")
		n_version = f"{ma}.{int(mi)+1}.0"
		break
	elif uinput == "p":
		print(f"Incrementing patch (p) version")
		n_version = f"{ma}.{mi}.{int(p)+1}"
		break
	elif uinput == "b":
		print(f"Build only (b) release, not incrementing version")
		break
	else:
		print(f"'{uinput}' is not a valid choice!")
		uinput = None

version_writer(n_version) # Save new version to file
version_printer(title = "----New Build Info----")

#* 2 - Pyinstaller
print(f"\nBuilding EXE with 'pyinstaller {spec_file}'")
subprocess.run(f"pyinstaller {spec_file}")

#* 3 - Rename EXE
n_exe_name = f"{og_exe_name}-{n_version}-{_version.__build__}"
print(f"\nRenaming new package '{og_exe_name}.exe' to '{n_exe_name}.exe'")
os.rename(os.path.join(dist_folder, og_exe_name+".exe"), os.path.join(dist_folder, n_exe_name+".exe"))
print("Packaging complete!")
sys.stdout.write('\a') # bell sound
sys.stdout.flush()
time.sleep(1)
sys.stdout.write('\a') # bell sound
sys.stdout.flush()
time.sleep(1)