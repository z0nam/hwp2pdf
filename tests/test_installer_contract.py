from pathlib import Path


INSTALLER = Path(__file__).resolve().parent.parent / "installer" / "hwp2pdf.iss"


def installer_source():
    return INSTALLER.read_text(encoding="utf-8")


def test_installer_offers_cli_path_task_and_notifies_windows():
    source = installer_source()

    assert "ChangesEnvironment=yes" in source
    assert 'Name: "addtopath"' in source
    assert 'Description: "{cm:AddToPath}"' in source
    assert 'Flags: unchecked' not in next(
        line for line in source.splitlines() if 'Name: "addtopath"' in line
    )


def test_installer_manages_only_its_own_path_token():
    source = installer_source()

    assert "function AddPathEntry" in source
    assert "function RemovePathEntry" in source
    assert "InstallerPathEntry" in source
    assert "RegDeleteValue(HKLM, AppRegistryKey, PathMarkerName)" in source
    assert "{olddata}" not in source
    assert "EnvironmentKey" in source
    assert "SplitString" not in source
    assert "FirstKept" in source
