from docxtpl import DocxTemplate
from pathlib import Path
import json
import subprocess
import sys

PACKAGE = "boto3"
DAYS = "30"

h = DocxTemplate("templates/docx_t.docx")
PYPINFO = Path(sys.exec_prefix, "bin/pypinfo")


def run_pypinfo(cmd: list) -> dict:
    output = subprocess.run(cmd, capture_output=True)
    print(output.stdout)
    print()
    return json.loads(output.stdout)


def get_python_version_for(pkg: str) -> dict:
    cmd = [PYPINFO, "-d", DAYS, "-j", pkg, "pyversion"]
    return run_pypinfo(cmd)


def get_countries_for(pkg: str) -> dict:
    cmd = [PYPINFO, "-d", DAYS, "-j", pkg, "country"]
    return run_pypinfo(cmd)


def get_platforms_for(pkg: str) -> dict:
    cmd = [PYPINFO, "-d", DAYS, "-j", pkg, "system", "distro"]
    return run_pypinfo(cmd)


def get_most_popular() -> dict:
    cmd = [PYPINFO, "-d", "365", "-j", "", "project"]
    print(cmd)
    return run_pypinfo(cmd)


py_version = get_python_version_for(PACKAGE)
download_count = sum([int(i["download_count"]) for i in py_version["rows"]])
country = get_countries_for(PACKAGE)
platform = get_platforms_for(PACKAGE)
popular = get_most_popular()

h.render(
    context={
        "package": PACKAGE,
        "download_count": download_count,
        "python_version": py_version,
        "country": country,
        "platforms": platform,
        "duration": DAYS,
        "popular": popular,
    }
)
h.save(filename="output_file.docx")
