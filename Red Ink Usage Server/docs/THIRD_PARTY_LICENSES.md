# Third-party component licenses — Red Ink Usage Server 1.3.0

The application uses the Python standard library, its own source and the **pinned** dependencies in `requirements.txt` and `requirements-build.txt`. There are no external JS/CSS libraries or hosted fonts in the dashboard. The included Red Ink logo is a first-party asset, not open-source stock artwork.

The declared runtime dependencies use permissive license families. The release build copies each **directly pinned** dependency's installed notice into `third-party-licenses/<package>/` and records package/version/reviewed license in a manifest. The build rejects unknown direct pins and missing notices. **This is not yet a complete automated audit of every transitively installed distribution or of the final EXE contents.** A Windows release therefore still requires a transitive dependency and packaged-file audit; do not treat the allowlist check alone as a legal approval. After dependency upgrades, repeat this review before changing the allowlist in `scripts/collect-third-party-notices.py`.

| Package(s) | Reviewed license | Source |
|---|---|---|
| Flask, Werkzeug, Jinja2, itsdangerous, Click, MarkupSafe | BSD-3-Clause | https://palletsprojects.com |
| Waitress | ZPL-2.1 (permissive) | https://github.com/Pylons/waitress |
| cheroot | BSD-3-Clause | https://github.com/cherrypy/cheroot |
| cryptography | Apache-2.0 OR BSD-3-Clause | https://github.com/pyca/cryptography |
| blinker | MIT | https://github.com/pallets-eco/blinker |
| more-itertools | MIT | https://github.com/more-itertools/more-itertools |
| jaraco.functools | MIT | https://github.com/jaraco/jaraco.functools |
| cffi | MIT | https://github.com/python-cffi/cffi |
| pycparser | BSD-3-Clause | https://github.com/eliben/pycparser |
| PyInstaller (build tool only) | **GPL-2.0-or-later WITH Bootloader-exception** (explicit exception) | https://pyinstaller.org/en/stable/license.html |

PyInstaller is the specifically permitted exception to the permissive-license rule. The PyInstaller license and bootloader exception allow distribution of applications bundled with it subject to the licenses of their runtime dependencies. No PyInstaller source is modified. Copying its notice file into releases is an additional transparency measure.

The included interpreter/runtime may have its own notices (including Python Software Foundation and platform/system components). Distributing binaries requires the builder to ensure these notices are shipped as appropriate, and the final release manifest must be checked. This list is a technical compliance aid and **does not replace** a final legal audit of a built release.

The separate Red Ink add-in repository has its own `license.txt` and dependencies; those should not be confused with the Python server's distribution.


## PyInstaller community-hooks qualification

PyInstaller depends on **`pyinstaller-hooks-contrib`**, whose *build-time standard hooks* are **GPL-2.0-or-later**, while its runtime hooks (those bundled with executables) are **Apache-2.0**. The PyInstaller bootloader exception explicitly permits commercial application bundling, but the community-hooks package is separately licensed; **it cannot be described without qualification as an entirely permissive build dependency**. See https://github.com/pyinstaller/pyinstaller-hooks-contrib/blob/master/LICENSE and https://pyinstaller.org/en/v6.22.3/license.html. The current packaging route uses these hooks as build tooling and does not deliberately distribute GPL standard-hook source files. If your policy bans use of *any* GPL build tooling other than PyInstaller itself, review/approve the hooks exception or choose another distribution route before shipping an EXE. This licensing qualification does not affect starting the source server directly with Python.

## URLs for redistribution notices (31 October 2026)

The PyInstaller notice and bootloader exception must be included with the frozen
Windows distribution. The release build copies the **installed version's**
license texts into `third-party-licenses/pyinstaller/`, and identifies the exact
package versions in `third-party-licenses/manifest.json`. Upstream references:

- https://pyinstaller.org/en/stable/license.html
- https://github.com/pyinstaller/pyinstaller
- https://github.com/pyinstaller/pyinstaller-hooks-contrib/blob/master/LICENSE

**Not applicable to Red Ink Usage Server:** The provided general `license.txt`
links `https://redink.ai/content/RedInk-pystray-LGPL-compliance.zip` to *pystray
in redink-helper*. The Usage Server does not bundle pystray, so publishing that
source/replacement package for the Usage Server would be misleading. PyInstaller's
bootloader exception does not require the unrelated helper's pystray package.
Check all transitive/bundled license notices against the exact Windows release
before distribution; a source-level allowlist is not final legal clearance.
