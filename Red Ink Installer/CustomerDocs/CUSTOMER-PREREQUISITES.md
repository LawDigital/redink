# Red Ink — Microsoft Prerequisites for Enterprise IT

Red Ink VSTO add-ins require Microsoft components that are normally present on current managed Office PCs, but organizations should detect them before deployment.

## 1. Microsoft Office

A supported desktop Microsoft Office installation is required. Deploy the Red Ink x86 MSI for 32-bit Office and the x64 MSI for 64-bit Office.

## 2. .NET Framework 4.8 or newer compatible 4.x runtime

Red Ink targets .NET Framework 4.8.

Official Microsoft download page:
https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48

For end-user machines choose the Runtime, not the Developer Pack.

.NET Framework 4.5 and later are an in-place update family; do not deploy .NET Framework 4.7.2 separately when 4.8 or a later compatible 4.x runtime is already present.

## 3. Microsoft Visual Studio 2010 Tools for Office Runtime

The VSTO Runtime is required for VSTO add-ins. It is often already installed with Office, but Microsoft provides the redistributable `vstor_redist.exe`.

Official Microsoft Download Center:
https://www.microsoft.com/en-us/download/details.aspx?id=105522

Microsoft Learn guidance:
https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime-installation-scenarios?view=visualstudio

## 4. Visual C++ Runtime for Word (if required)

The current Red Ink for Word ClickOnce project declares the Visual C++ 14 x64 runtime as an installation prerequisite. Until Red Ink confirms that the native dependency requiring it has been removed, deploy/detect the supported Microsoft Visual C++ v14 x64 Redistributable for Word installations.

Microsoft guidance and current download links:
https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-170

## Enterprise deployment options

Organizations may deploy prerequisites using their normal software distribution platform before Red Ink. This is preferred for centrally managed devices.

Examples:

- Microsoft Intune
- Microsoft Configuration Manager / SCCM
- Group Policy software deployment
- enterprise RMM/application management tools

For unmanaged/manual devices Red Ink may provide an optional Microsoft-generated Setup.exe bootstrapper configured to download prerequisite installers from the component vendor's web site.

## Detection helper

Run:

`powershell.exe -ExecutionPolicy Bypass -File .\Detect-RedInk-Prerequisites.ps1`

The script only reports status and does not install software.
