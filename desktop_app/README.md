# Split Data Pro — Windows Desktop

Windows 7 SP1 / 10 / 11 **32-bit portable EXE target**. Phase 1 preview build.

## Download the EXE (GitHub Actions)

1. Open the repository's **Actions** tab.
2. Select **Build Windows EXE (legacy compatible)**.
3. Click **Run workflow**, choosing `feature/windows-desktop-phase1`.
4. After the run succeeds, open its run page and download **SplitDataPro-Windows7-10-11-x86** under **Artifacts**.
5. Unzip and launch `SplitDataPro.exe`.

GitHub sign-in may be required to download artifacts. This is an automatic build recipe, **not** a claim that the EXE has already been built or tested on Windows 7.

Windows 7 SP1 with appropriate runtime updates may be required. Microsoft no longer supports Windows 7; use only on trusted offline files. Test on a Windows 7 PC before relying on compatibility.

## Run from source

Install **Python 3.8 32-bit** on Windows, then:

```powershell
cd desktop_app
py -3.8-32 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

Phase 1: sidebar dashboard, .xlsx import and 20-row preview, split into multiple .xlsx workbooks, progress and errors. Combine, Cleanup and File Tools are placeholders. VBA sources in the root remain untouched. Formatting and VBA macros in input files are not preserved in split output. Output is local, no login needed.
