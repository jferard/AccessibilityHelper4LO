# -*- coding: utf-8 -*-
#  Py4LO Extension Template - a template for Python extensions for
#  LibreOffice and a script to pack and install the extension.
#
#     Copyright (C) 2023 J. Férard <https://github.com/jferard>
#
#     This file is part of Py4LO Extension Template.
#
#     Py4LO Extension Template is free software: you can redistribute it and/or
#     modify it under the terms of the GNU General Public License as published
#     by the Free Software Foundation, either version 3 of the License, or
#     (at your option) any later version.
#
#     Py4LO Extension Template is distributed in the hope that it will be
#     useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <http://www.gnu.org/licenses/>.
import argparse
import glob
import os
import platform
import shutil
import subprocess
import sys
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from pprint import pprint

system = platform.system()
if system == "Windows":
    USER_PATH = Path.home() / r"AppData\Roaming\LibreOffice\4\user"
    PROGRAM_PATH = Path(r"C:\Program Files\LibreOffice 7\program")
    UNO_PKG_PATH = PROGRAM_PATH / "unopkg.com"
elif system == "Linux":
    USER_PATH = Path.home() / ".config/libreoffice/4/user"
    PROGRAM_PATH = Path("/usr/lib/libreoffice/program")
    UNO_PKG_PATH = PROGRAM_PATH / "unopkg"
else:
    raise Exception("You'll have to configure py4lo_extension_builder")

CACHE_PATH = USER_PATH / "uno_packages/cache"


class ExtensionPacker:
    def __init__(self, scr_path, temp_path):
        self.src_path = src_path
        self.temp_path = temp_path

    def install_extension(self):

        oxt_name = self.create_extension()
        self._install_extension(oxt_name)

    def create_extension(self):
        name = self._find_name()
        if name is None:
            print("Can't find name !! Please check description.xml file...")
        version = self._find_version()
        if version is None:
            print("Can't find version !! Please check description.xml file...")
        self._clean_up()
        oxt_name = self._create_extension(name, version)
        return oxt_name

    def _find_name(self) -> str:
        root = ET.parse(self.src_path / "description.xml").getroot()
        name = root.find(
            "{http://openoffice.org/extensions/description/2006}display-name/{http://openoffice.org/extensions/description/2006}name"
        ).text
        return name

    def _find_version(self) -> str:
        root = ET.parse(self.src_path / "description.xml").getroot()
        version = root.find(
            "{http://openoffice.org/extensions/description/2006}version").attrib[
            "value"]
        return version

    def _clean_up(self):
        print("> clean up")
        for lo_ext in glob.glob("*.oxt"):
            os.remove(lo_ext)

    def _create_extension(self, name: str, version: str) -> str:
        creator = ExtensionCreator(self.src_path, self.temp_path)
        return creator.create_extension(name, version)

    def _install_extension(self, oxt_name: str):
        arr = [
            str(UNO_PKG_PATH), "add", "-v", "-s", "-f", oxt_name
        ]
        print("> install", " ".join(arr))
        result = subprocess.run(
            arr, capture_output=True, text=True)
        print(">> out")
        print(result.stdout)
        if result.returncode or result.stderr.strip():
            print(">> err")
            print(result.stderr)
        print(result)
        print("See: {}".format(CACHE_PATH))


class ExtensionCreator:
    def __init__(self, src_path: Path, temp_path: Path):
        self._src_path = src_path
        self._temp_path = temp_path

    def create_extension(self, name: str, version: str) -> str:
        print(">> create extension")
        base_name = "{}-{}".format(name, version)
        zip_name = base_name + ".zip"
        oxt_name = base_name + ".oxt"
        #    txt_name = base_name + ".txt"
        shutil.rmtree(self._temp_path, ignore_errors=True)
        self._temp_path.mkdir()
        for filename in ("Addons.xcu", "description.xml"):
            shutil.copyfile(self._src_path / filename,
                            self._temp_path / filename)
        for p in self._src_path.glob("*.py"):
            shutil.copyfile(p, self._temp_path / p.name)
        self._copy_tree("description", "**/*")
        self._copy_tree("icons", "*.png")
        self._copy_tree("META-INF", "*.xml")
        self._copy_tree("pythonpath", "**/*.py")
        self._copy_tree("registration", "*.txt")

        self._make_archive(base_name)
        os.rename(zip_name, oxt_name)
        shutil.rmtree(self._temp_path, ignore_errors=True)
        return oxt_name

    def _copy_tree(self, subdir_name: str, pattern: str):
        src_subpath = self._src_path / subdir_name
        temp_subpath = self._temp_path / subdir_name
        temp_subpath.mkdir()
        for p in src_subpath.glob(pattern):
            shutil.copyfile(p, self._temp_path / p.relative_to(self._src_path))

    def _make_archive(self, base_name: str):
        import zipfile
        zip_filename = base_name + ".zip"
        archive_dir = os.path.dirname(base_name)
        if archive_dir and not os.path.exists(archive_dir):
            os.makedirs(archive_dir)
        with zipfile.ZipFile(zip_filename, "w",
                             compression=zipfile.ZIP_STORED) as zf:
            for p in self._temp_path.rglob("*"):
                if p.is_dir():
                    print("    * {}/".format(p.relative_to(self._temp_path)))
                    zf.write(p, p.relative_to(self._temp_path))
                else:
                    print("    * {}".format(p.relative_to(self._temp_path)))
                    zf.write(p, p.relative_to(self._temp_path))


class ExtensionTester:
    def test_extension(self):
        import flake8.main.application
        import pytest
        sys.path.insert(0, "src/pythonpath")

        retcode = pytest.main(
            ["--cov-report", "term-missing", "--cov=src/pythonpath"])
        if retcode == 0:
            retcode & pytest.main(
                ["--ignore=src\{}.py", "--cov-report", "term-missing",
                 "--cov-append", "--doctest-modules", ".",
                 "--cov=src/pythonpath"])
            if retcode == 0:
                app = flake8.main.application.Application()
                app.run(["src/pythonpath"])
                if not app.catastrophic_failure and app.result_count == 0:
                    return
        sys.exit(retcode)


class ExtensionsCleaner:
    def __init__(self, cache_path: Path, dry_run=True):
        self._cache_path = cache_path
        if dry_run:
            self._fs = DryRunFS()
        else:
            raise Exception()

    def clean_extension(self):
        self._rename(self._cache_path / "registry")
        self._rename(self._cache_path / "uno_packages.pmap")
        for p in (self._cache_path / "uno_packages").rglob("*"):
            if p.is_file():
                self._fs.unlink(p)
            elif p.is_dir():
                self._fs.rmtree(p)

    def _rename(self, path: Path):
        i = 0
        while True:
            i += 1
            new_path = path.parent / (path.stem + str(i) + path.suffix)
            if not new_path.exists():
                break

        self._fs.rename(path, new_path)


class DryRunFS:
    def rmtree(self, p: Path):
        print("rmtree {}:".format(p))
        pprint(list(p.rglob("*")))

    def unlink(self, p: Path):
        print("unlink {}".format(p))

    def rename(self, path: Path, new_path: Path):
        print("rename {} to {}".format(path, new_path))


class DangerFS:
    def rmtree(self, p: Path):
        shutil.rmtree(p)

    def unlink(self, p: Path):
        p.unlink()

    def rename(self, path: Path, new_path: Path):
        path.rename(new_path)


def prepare():
    base_url = "https://raw.githubusercontent.com/jferard/py4lo/master/lib/"
    for file in ("py4lo_typing.py", "py4lo_helper.py", "py4lo_commons.py",
                 "py4lo_dialogs.py", "mock_constants.py"):
        url = base_url + file
        print("Try to download", url)
        urllib.request.urlretrieve(url, "src/pythonpath/" + file)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog='py4lo_extension_builder',
        description='Builds an extension')
    parser.add_argument("command")
    parser.add_argument('-n', '--nodryrun',
                        action='store_true', default=False)  # on/off flag
    args = parser.parse_args()
    # TODO config
    src_path = Path("src")
    temp_path = Path("temp")
    if args.command == "prepare":
        prepare()
    elif args.command == "install":
        prepare()
        ExtensionTester().test_extension()
        ExtensionPacker(src_path, temp_path).install_extension()
    elif args.command == "create":
        prepare()
        ExtensionPacker(src_path, temp_path).create_extension()
    elif args.command == "test":
        prepare()
        ExtensionTester().test_extension()
    elif args.command == "clean":
        dry_run = not args.nodryrun
        if not dry_run:
            ret = input("Are you sure you want to remove all extensions"
                        " from your installation? [y/N]")
            if ret.casefold() not in ("y", "yes"):
                sys.exit(0)
        else:
            print("DRY RUN")
        ExtensionsCleaner(CACHE_PATH, dry_run).clean_extension()
