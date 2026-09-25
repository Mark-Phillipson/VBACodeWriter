GitHub Copilot

Steps (concise) to create an MSI using the Visual Studio Installer Projects extension

1. Install the extension
- In Visual Studio: Extensions > Manage Extensions > search “Visual Studio Installer Projects” (Microsoft) > Download > Restart Visual Studio.

2. Add a Setup project to the solution
- Right-click the solution > Add > New Project.
- Choose “Setup Project” (under “Other Project Types” → “Setup and Deployment” or “Visual Studio Installer”), give it a name.

3. Add your app outputs and files
- Open the Setup project and select File System on Target Machine.
- Right-click Application Folder > Add > Project Output > choose your project > select “Primary output” (and “Content Files” if needed).
- Add any extra files (config, icons) by Add > File.

4. COM interop / Office interop registration
- If your assembly exposes COM types:
  - Preferred: In your class library project Properties → Build, enable “Register for COM interop” for development only.
  - In the Setup project File System view, select the assembly (Primary output) and set its Properties → Register to the COM registration option (so the installer registers COM classes on install).
- Note: Installer must run elevated to register COM entries (HKCR). If you prefer, you can call regasm in a custom action, but MSI COM registration is more robust.

5. Prerequisites and launch conditions
- Setup Project → Properties → Prerequisites: include .NET Framework, VC++ runtimes, etc.
- Use Launch Conditions editor to require Office/Access if your app depends on it (check registry or files).

6. Custom actions (optional)
- If you have installer classes or need custom steps, open Custom Actions view and add the Primary output (or specific EXE) to Install / Commit / Rollback / Uninstall as required.

7. UI and shortcuts
- Edit User Interface to add dialogs (license, readme, install folder).
- In File System view create shortcuts in the User’s Programs Menu or Desktop.

8. Install scope and location
- In Setup project Properties set DefaultLocation (ProgramFilesFolder vs LocalAppData) and set InstallAllUsers appropriately for machine-wide installs.

9. Build the installer
- In Visual Studio: right-click Setup project → Build. The MSI (and setup.exe if prerequisites selected) appears in the Setup project’s Release (or Debug) output folder.
- Command-line CI: use devenv.exe Solution.sln /Build "Release" (vdproj files are not built by msbuild). Use Visual Studio build agent or call devenv in your pipeline.

10. Test the installer
- Test on a clean VM or test machine.
- Run the MSI as Administrator (if registering COM). Verify COM registration, Office integration, Add/Remove Programs entry, uninstall behavior.

11. Signing and distribution
- Sign the MSI (recommended) with signtool or configure in your release pipeline.
- For internal distribution, consider packaging with Chocolatey/Winget or deploying via IT tools.

Notes / gotchas
- COM/Office interop requires admin elevation for per-machine registration. If you want per-user installs without registry writes, you’ll need different registration strategies.
- vdproj installers are easy to author but less flexible than WiX for advanced scenarios (upgrade rules, complex registry logic).
- For automated builds in CI, ensure a build machine with Visual Studio installed so devenv can build the installer.

If you want, I can:
- Add a template Setup project to your repo that includes primary output, COM registration property set, and prerequisites configured; or
- Create a short checklist for testing COM registration and Access interop on a VM. Which would you like?