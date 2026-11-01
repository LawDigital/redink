"""Static release safety gate; Windows execution is a separate required build step."""
import unittest
from pathlib import Path


class WindowsBuildContractTests(unittest.TestCase):
    def test_build_refuses_deployment_state_before_output_cleanup(self):
        source = (Path(__file__).resolve().parents[1] / "build-windows.ps1").read_text(encoding="utf-8")
        guard_call = "Assert-NoDeploymentData -Directory $OutputRoot"
        cleanup = "Remove-Item -LiteralPath $OutputRoot -Recurse -Force"
        self.assertEqual(source.count(cleanup), 1)
        self.assertIn(guard_call, source)
        self.assertLess(source.index(guard_call), source.index(cleanup))
        guard = source.split("function Assert-NoDeploymentData", 1)[1].split("$previousMode", 1)[0]
        self.assertIn("-Recurse -Force -ErrorAction Stop", guard)
        for marker in ("'data'", "'secrets.json'", "'server.toml'", "'server.lock'",
                       "'*.sqlite3'", "'*.sqlite3-wal'", "'*.sqlite3-shm'"):
            self.assertIn(marker, guard)
        self.assertIn("throw ", guard)
        self.assertIn("Nothing was deleted.", guard)


    def test_vischer_certificate_is_default_and_unsigned_build_requires_opt_in(self):
        source = (Path(__file__).resolve().parents[1] / 'build-windows.ps1').read_text(encoding='utf-8')
        self.assertIn("[string]$CertificateSubject = 'VISCHER AG'", source)
        self.assertIn('Select-CodeSigningCertificate -Owner $CertificateSubject', source)
        self.assertIn('1.3.6.1.5.5.7.3.3',source)
        self.assertIn('HasPrivateKey', source)
        self.assertIn('$SkipSigning',source)
        self.assertIn('if (-not $SkipSigning)',source)
        self.assertIn("Invoke-Checked $Tool @('verify'",source)
        self.assertIn("'signed-and-verified'",source)

if __name__ == "__main__":
    unittest.main()
