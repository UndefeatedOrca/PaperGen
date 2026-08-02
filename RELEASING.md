# Releasing PaperGen

## What belongs in Git

Commit the source code, templates, installer script, PyInstaller spec, build script, and documentation:

- `papergen.py`
- `templategen.py`
- `templates/`
- `PaperGen.spec`
- `installer/PaperGen.iss`
- `build.bat`
- `requirements-build.txt`
- `readme.md`
- `LICENSE`

Do not commit `.venv-build/`, `build/`, `dist/`, or local JSON data. They are ignored by `.gitignore`.

## Create a release

1. Update `MyAppVersion` in `installer/PaperGen.iss`. Use a version such as `0.1.0`, without the leading `v`.

2. Review the source changes and run a clean build from the repository root:

   ```powershell
   build.bat
   ```

3. Test `dist\installer\PaperGenSetup.exe` on a second Windows account or clean test machine. Verify configuration, templates, document generation, the Explorer right-click menu, upgrades, and uninstall behavior.

4. Commit the release changes:

   ```powershell
   git add .
   git commit -m "Prepare PaperGen v0.1.0 release"
   ```

5. Create and push an annotated tag:

   ```powershell
   git tag -a v0.1.0 -m "PaperGen v0.1.0"
   git push origin main
   git push origin v0.1.0
   ```

6. On GitHub, open the repository's **Releases** page, choose **Draft a new release**, select the `v0.1.0` tag, and attach:

   ```text
   dist\installer\PaperGenSetup.exe
   ```

7. Add release notes summarizing user-visible changes, then publish the release.

GitHub's automatically generated source archive is useful for developers. The `PaperGenSetup.exe` attachment is the file ordinary Windows users should download.

## Versioning convention

Use semantic versions:

- Patch: bug fixes, such as `0.1.1`
- Minor: backward-compatible features, such as `0.2.0`
- Major: breaking behavior or configuration changes, such as `1.0.0`

Keep the Inno Setup `MyAppVersion` value and Git tag synchronized.
