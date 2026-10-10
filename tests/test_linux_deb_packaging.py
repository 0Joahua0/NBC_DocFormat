"""Build and inspect a real Debian archive with a small stand-in bundle."""

from pathlib import Path
import shutil
import subprocess
import sys

import pytest


PROJECT_ROOT = Path(__file__).resolve().parent.parent


@pytest.mark.skipif(sys.platform != 'linux' or not shutil.which('dpkg-deb'),
                    reason='Requires Linux and dpkg-deb')
@pytest.mark.parametrize('arch', ['amd64', 'arm64'])
def test_deb_installs_bundle_desktop_icon_and_preserves_modes(tmp_path, arch):
    repo = tmp_path / 'project with spaces'
    script_dir = repo / 'packaging' / 'linux'
    script_dir.mkdir(parents=True)
    script = script_dir / 'build-deb.sh'
    shutil.copy2(PROJECT_ROOT / 'packaging/linux/build-deb.sh', script)
    for name in ['packaging/appimage/NBC_DocFormat.desktop', 'assets/icon.png',
                 'LICENSE', 'LICENSE-HISTORY.md']:
        target = repo / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(PROJECT_ROOT / name, target)
    (repo / 'build.py').write_text('VERSION = "1.0.3"\n')
    bundle = repo / 'bundle'
    internal = bundle / '_internal'
    internal.mkdir(parents=True)
    executable = bundle / 'NBC_DocFormat_linux'
    executable.write_text('#!/bin/sh\nexit 0\n')
    executable.chmod(0o755)
    (internal / 'library.so.1').write_bytes(b'library fixture')
    (internal / 'library.so').symlink_to('library.so.1')

    subprocess.run(['bash', str(script), str(bundle), arch], check=True)
    package = repo / 'dist' / f'NBC_DocFormat_linux_{arch}.deb'
    fields = subprocess.check_output(['dpkg-deb', '-f', str(package)], text=True)
    assert f'Architecture: {arch}' in fields
    assert 'Package: nbc-docformat' in fields
    assert 'Version: 1.0.3' in fields
    listing = subprocess.check_output(['dpkg-deb', '-c', str(package)], text=True)
    assert 'root/root' in listing
    extracted = tmp_path / 'extracted'
    subprocess.run(['dpkg-deb', '-R', str(package), str(extracted)], check=True)
    app_dir = extracted / 'opt/nbc-docformat'
    assert (app_dir / 'NBC_DocFormat_linux').stat().st_mode & 0o111
    assert (app_dir / '_internal/library.so').read_bytes() == b'library fixture'
    launcher = extracted / 'usr/bin/NBC_DocFormat_linux'
    assert launcher.readlink() == Path('/opt/nbc-docformat/NBC_DocFormat_linux')
    desktop = extracted / 'usr/share/applications/NBC_DocFormat.desktop'
    if shutil.which('desktop-file-validate'):
        subprocess.run(['desktop-file-validate', str(desktop)], check=True)
    assert 'StartupWMClass=Nbc_docformat' in desktop.read_text()
    assert (extracted / 'usr/share/icons/hicolor/256x256/apps/NBC_DocFormat.png').exists()
    for hook in ['postinst', 'postrm']:
        path = extracted / 'DEBIAN' / hook
        assert path.stat().st_mode & 0o111
        subprocess.run(['sh', '-n', str(path)], check=True)
