"""
Functional tests for the PowerPointFilter implementation in DSpace 11.

These tests verify:
1. Fixture files (test.ppt, test.pptx, test.ppt.txt, test.pptx.txt) exist with expected content
2. dspace.cfg correctly registers PowerPointFilter and excludes PPT/PPTX from TikaTextExtractionFilter
3. PowerPointFilter.java class exists with expected structure
4. Maven unit tests (PowerPointFilterTest, 11 tests) all pass
"""

import subprocess
import os
import pytest
import xml.etree.ElementTree as ET
import time

WORKSPACE_DIR = os.environ.get("WORKSPACE_DIR", "/l2l/workspace")
DSPACE_DIR = os.path.join(WORKSPACE_DIR, "DSpace")
MEDIAFILTER_RESOURCES = os.path.join(
    DSPACE_DIR,
    "dspace-api/src/test/resources/org/dspace/app/mediafilter"
)
MEDIAFILTER_SRC = os.path.join(
    DSPACE_DIR,
    "dspace-api/src/main/java/org/dspace/app/mediafilter"
)
DSPACE_CFG = os.path.join(DSPACE_DIR, "dspace/config/dspace.cfg")
JAVA_HOME = os.environ.get("JAVA_HOME", "/usr/lib/jvm/temurin-21-jdk-amd64")
MAVEN_HOME = os.environ.get("MAVEN_HOME", "/opt/maven")
MVN = os.path.join(MAVEN_HOME, "bin", "mvn")


def run_maven(*args, timeout=600):
    """Run a Maven command in the DSpace directory."""
    env = os.environ.copy()
    env["JAVA_HOME"] = JAVA_HOME
    env["MAVEN_HOME"] = MAVEN_HOME
    env["PATH"] = f"{MAVEN_HOME}/bin:{JAVA_HOME}/bin:{env.get('PATH', '')}"
    cmd = [MVN] + list(args)
    result = subprocess.run(
        cmd,
        cwd=DSPACE_DIR,
        capture_output=True,
        text=True,
        timeout=timeout,
        env=env,
    )
    return result


class TestFixtureFiles:
    """Verify that fixture files exist and contain required content."""

    def test_ppt_fixture_exists(self):
        """test.ppt fixture file must be present."""
        path = os.path.join(MEDIAFILTER_RESOURCES, "test.ppt")
        assert os.path.isfile(path), f"test.ppt not found at {path}"
        assert os.path.getsize(path) > 0, "test.ppt must not be empty"

    def test_pptx_fixture_exists(self):
        """test.pptx fixture file must be present."""
        path = os.path.join(MEDIAFILTER_RESOURCES, "test.pptx")
        assert os.path.isfile(path), f"test.pptx not found at {path}"
        assert os.path.getsize(path) > 0, "test.pptx must not be empty"

    def test_ppt_expected_output_exists(self):
        """test.ppt.txt expected-output fixture must be present."""
        path = os.path.join(MEDIAFILTER_RESOURCES, "test.ppt.txt")
        assert os.path.isfile(path), f"test.ppt.txt not found at {path}"

    def test_pptx_expected_output_exists(self):
        """test.pptx.txt expected-output fixture must be present."""
        path = os.path.join(MEDIAFILTER_RESOURCES, "test.pptx.txt")
        assert os.path.isfile(path), f"test.pptx.txt not found at {path}"

    def test_ppt_expected_output_contains_quick_brown_fox(self):
        """test.ppt.txt must contain 'quick brown fox' (DSpace 6 parity phrase)."""
        path = os.path.join(MEDIAFILTER_RESOURCES, "test.ppt.txt")
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "quick brown fox" in content.lower(), (
            f"test.ppt.txt does not contain 'quick brown fox'. Content: {content!r}"
        )

    def test_pptx_expected_output_contains_quick_brown_fox(self):
        """test.pptx.txt must contain 'quick brown fox' (DSpace 6 parity phrase)."""
        path = os.path.join(MEDIAFILTER_RESOURCES, "test.pptx.txt")
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "quick brown fox" in content.lower(), (
            f"test.pptx.txt does not contain 'quick brown fox'. Content: {content!r}"
        )

    def test_ppt_expected_output_contains_dspace_presentation_phrase(self):
        """test.ppt.txt must contain DSpace presentation phrase."""
        path = os.path.join(MEDIAFILTER_RESOURCES, "test.ppt.txt")
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "DSpace can extract the text from your presentations!" in content, (
            f"test.ppt.txt missing expected DSpace phrase. Content: {content!r}"
        )

    def test_pptx_expected_output_contains_dspace_presentation_phrase(self):
        """test.pptx.txt must contain DSpace presentation phrase."""
        path = os.path.join(MEDIAFILTER_RESOURCES, "test.pptx.txt")
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        assert "DSpace can extract the text from your presentations!" in content, (
            f"test.pptx.txt missing expected DSpace phrase. Content: {content!r}"
        )


class TestDSpaceConfiguration:
    """Verify dspace.cfg registers PowerPointFilter correctly."""

    def _read_cfg(self):
        with open(DSPACE_CFG, "r", encoding="utf-8") as f:
            return f.read()

    def test_powerpoint_filter_registered_as_named_plugin(self):
        """PowerPointFilter must be registered as a named FormatFilter plugin."""
        content = self._read_cfg()
        assert "org.dspace.app.mediafilter.PowerPointFilter" in content, (
            "PowerPointFilter is not registered as a FormatFilter plugin in dspace.cfg"
        )

    def test_powerpoint_filter_named_plugin_has_correct_name(self):
        """PowerPointFilter plugin entry must map to 'PowerPoint Text Extractor'."""
        content = self._read_cfg()
        assert "PowerPoint Text Extractor" in content, (
            "PowerPointFilter not registered with name 'PowerPoint Text Extractor' in dspace.cfg"
        )

    def test_powerpoint_filter_input_formats_ppt(self):
        """PowerPointFilter.inputFormats must include 'Microsoft Powerpoint'."""
        content = self._read_cfg()
        assert "filter.org.dspace.app.mediafilter.PowerPointFilter.inputFormats = Microsoft Powerpoint" in content, (
            "'Microsoft Powerpoint' missing from PowerPointFilter.inputFormats in dspace.cfg"
        )

    def test_powerpoint_filter_input_formats_pptx(self):
        """PowerPointFilter.inputFormats must include 'Microsoft Powerpoint XML'."""
        content = self._read_cfg()
        assert "filter.org.dspace.app.mediafilter.PowerPointFilter.inputFormats = Microsoft Powerpoint XML" in content, (
            "'Microsoft Powerpoint XML' missing from PowerPointFilter.inputFormats in dspace.cfg"
        )

    def test_tika_filter_does_not_include_ppt(self):
        """TikaTextExtractionFilter must NOT list 'Microsoft Powerpoint' as inputFormat (exclusive ownership)."""
        content = self._read_cfg()
        # Find all inputFormats lines for TikaTextExtractionFilter
        tika_lines = [
            line for line in content.splitlines()
            if "filter.org.dspace.app.mediafilter.TikaTextExtractionFilter.inputFormats" in line
        ]
        ppt_in_tika = any("Microsoft Powerpoint" in line and "XML" not in line for line in tika_lines)
        assert not ppt_in_tika, (
            "'Microsoft Powerpoint' must be removed from TikaTextExtractionFilter.inputFormats "
            f"to give PowerPointFilter exclusive ownership. Tika lines: {tika_lines}"
        )

    def test_tika_filter_does_not_include_pptx(self):
        """TikaTextExtractionFilter must NOT list 'Microsoft Powerpoint XML' as inputFormat."""
        content = self._read_cfg()
        tika_lines = [
            line for line in content.splitlines()
            if "filter.org.dspace.app.mediafilter.TikaTextExtractionFilter.inputFormats" in line
        ]
        pptx_in_tika = any("Microsoft Powerpoint XML" in line for line in tika_lines)
        assert not pptx_in_tika, (
            "'Microsoft Powerpoint XML' must be removed from TikaTextExtractionFilter.inputFormats "
            f"to give PowerPointFilter exclusive ownership. Tika lines: {tika_lines}"
        )


class TestPowerPointFilterClass:
    """Verify the PowerPointFilter.java class structure."""

    def _read_source(self):
        path = os.path.join(MEDIAFILTER_SRC, "PowerPointFilter.java")
        assert os.path.isfile(path), f"PowerPointFilter.java not found at {path}"
        with open(path, "r", encoding="utf-8") as f:
            return f.read()

    def test_class_file_exists(self):
        """PowerPointFilter.java must exist in the mediafilter package."""
        path = os.path.join(MEDIAFILTER_SRC, "PowerPointFilter.java")
        assert os.path.isfile(path), f"PowerPointFilter.java not found at {path}"

    def test_extends_media_filter(self):
        """PowerPointFilter must extend MediaFilter."""
        content = self._read_source()
        assert "extends MediaFilter" in content, (
            "PowerPointFilter must extend MediaFilter"
        )

    def test_implements_self_register_input_formats(self):
        """PowerPointFilter must implement SelfRegisterInputFormats."""
        content = self._read_source()
        assert "implements SelfRegisterInputFormats" in content, (
            "PowerPointFilter must implement SelfRegisterInputFormats"
        )

    def test_returns_text_bundle(self):
        """PowerPointFilter.getBundleName() must return 'TEXT'."""
        content = self._read_source()
        assert '"TEXT"' in content, (
            "PowerPointFilter.getBundleName() must return 'TEXT'"
        )

    def test_uses_poi_extractor_factory(self):
        """PowerPointFilter must use Apache POI ExtractorFactory for direct extraction."""
        content = self._read_source()
        assert "ExtractorFactory" in content, (
            "PowerPointFilter must use Apache POI ExtractorFactory for direct extraction"
        )

    def test_handles_null_source(self):
        """PowerPointFilter must return null for null source (per FormatFilter contract)."""
        content = self._read_source()
        assert "source == null" in content or "if (source == null)" in content, (
            "PowerPointFilter must explicitly handle null source and return null"
        )

    def test_sets_slides_by_default(self):
        """PowerPointFilter must call setSlidesByDefault(true) for behavioral parity."""
        content = self._read_source()
        assert "setSlidesByDefault(true)" in content, (
            "PowerPointFilter must call setSlidesByDefault(true)"
        )

    def test_sets_notes_by_default(self):
        """PowerPointFilter must call setNotesByDefault(true) for behavioral parity."""
        content = self._read_source()
        assert "setNotesByDefault(true)" in content, (
            "PowerPointFilter must call setNotesByDefault(true)"
        )

    def test_mime_types_include_ppt(self):
        """PowerPointFilter must declare application/vnd.ms-powerpoint MIME type."""
        content = self._read_source()
        assert "application/vnd.ms-powerpoint" in content, (
            "PowerPointFilter must declare application/vnd.ms-powerpoint in getInputMIMETypes()"
        )

    def test_mime_types_include_pptx(self):
        """PowerPointFilter must declare PPTX MIME type."""
        content = self._read_source()
        assert "application/vnd.openxmlformats-officedocument.presentationml.presentation" in content, (
            "PowerPointFilter must declare .pptx MIME type in getInputMIMETypes()"
        )


class TestMavenUnitTests:
    """Run the Maven unit tests for PowerPointFilterTest."""

    @pytest.fixture(scope="class")
    def maven_test_result(self):
        """Run mvn test for PowerPointFilterTest and return the result."""
        env = os.environ.copy()
        env["JAVA_HOME"] = JAVA_HOME
        env["MAVEN_HOME"] = MAVEN_HOME
        env["PATH"] = f"{MAVEN_HOME}/bin:{JAVA_HOME}/bin:{env.get('PATH', '')}"

        # Copy solr.xml if needed (required by EmbeddedSolrClientFactory)
        solr_xml_src = os.path.join(DSPACE_DIR, "dspace-api/src/test/data/solr/solr.xml")
        solr_xml_dst = os.path.join(DSPACE_DIR, "dspace-api/target/testing/dspace/solr/solr.xml")
        if os.path.isfile(solr_xml_src) and not os.path.isfile(solr_xml_dst):
            os.makedirs(os.path.dirname(solr_xml_dst), exist_ok=True)
            import shutil
            shutil.copy2(solr_xml_src, solr_xml_dst)

        start = time.time()
        result = subprocess.run(
            [
                MVN, "test",
                "-pl", "dspace-api",
                "-P", "!test-environment",
                "--no-transfer-progress",
                "-DskipUnitTests=false",
                "-DskipIntegrationTests=true",
                "-Dtest=PowerPointFilterTest",
                "-Dsurefire.failIfNoSpecifiedTests=false",
            ],
            cwd=DSPACE_DIR,
            capture_output=True,
            text=True,
            timeout=600,
            env=env,
        )
        result.duration_ms = (time.time() - start) * 1000
        return result

    def test_maven_exits_zero(self, maven_test_result):
        """Maven test command must exit with code 0 (all tests pass)."""
        assert maven_test_result.returncode == 0, (
            f"Maven test failed with exit code {maven_test_result.returncode}.\n"
            f"STDOUT (last 3000 chars):\n{maven_test_result.stdout[-3000:]}\n"
            f"STDERR (last 1000 chars):\n{maven_test_result.stderr[-1000:]}"
        )

    def test_surefire_report_exists(self, maven_test_result):
        """Surefire report must exist after test run."""
        report_path = os.path.join(
            DSPACE_DIR,
            "dspace-api/target/surefire-reports/org.dspace.app.mediafilter.PowerPointFilterTest.txt"
        )
        assert os.path.isfile(report_path), (
            f"Surefire report not found at {report_path}. "
            f"Maven stdout (last 2000 chars): {maven_test_result.stdout[-2000:]}"
        )

    def test_all_eleven_tests_ran(self, maven_test_result):
        """All 11 PowerPointFilterTest tests must run."""
        report_path = os.path.join(
            DSPACE_DIR,
            "dspace-api/target/surefire-reports/org.dspace.app.mediafilter.PowerPointFilterTest.txt"
        )
        if not os.path.isfile(report_path):
            pytest.skip("Surefire report not found")

        with open(report_path, "r") as f:
            content = f.read()

        # Parse: "Tests run: 11, Failures: 0, Errors: 0, Skipped: 0"
        import re
        match = re.search(r"Tests run:\s*(\d+)", content)
        assert match, f"Could not parse 'Tests run' from surefire report: {content}"
        tests_run = int(match.group(1))
        assert tests_run == 11, (
            f"Expected 11 tests to run, but got {tests_run}. Report: {content}"
        )

    def test_zero_failures(self, maven_test_result):
        """PowerPointFilterTest must have 0 failures."""
        report_path = os.path.join(
            DSPACE_DIR,
            "dspace-api/target/surefire-reports/org.dspace.app.mediafilter.PowerPointFilterTest.txt"
        )
        if not os.path.isfile(report_path):
            pytest.skip("Surefire report not found")

        with open(report_path, "r") as f:
            content = f.read()

        import re
        match = re.search(r"Failures:\s*(\d+)", content)
        assert match, f"Could not parse 'Failures' from surefire report: {content}"
        failures = int(match.group(1))
        assert failures == 0, (
            f"Expected 0 failures, but got {failures}. Report:\n{content}"
        )

    def test_zero_errors(self, maven_test_result):
        """PowerPointFilterTest must have 0 errors."""
        report_path = os.path.join(
            DSPACE_DIR,
            "dspace-api/target/surefire-reports/org.dspace.app.mediafilter.PowerPointFilterTest.txt"
        )
        if not os.path.isfile(report_path):
            pytest.skip("Surefire report not found")

        with open(report_path, "r") as f:
            content = f.read()

        import re
        match = re.search(r"Errors:\s*(\d+)", content)
        assert match, f"Could not parse 'Errors' from surefire report: {content}"
        errors = int(match.group(1))
        assert errors == 0, (
            f"Expected 0 errors, but got {errors}. Report:\n{content}"
        )

    def test_build_success_in_stdout(self, maven_test_result):
        """Maven output must contain BUILD SUCCESS."""
        assert "BUILD SUCCESS" in maven_test_result.stdout, (
            f"BUILD SUCCESS not found in Maven output.\n"
            f"Maven stdout (last 3000 chars):\n{maven_test_result.stdout[-3000:]}"
        )

    def test_parity_tests_mentioned_in_output(self, maven_test_result):
        """Maven output should mention parity tests ran."""
        stdout = maven_test_result.stdout
        # Look for test class running in surefire output
        assert "PowerPointFilterTest" in stdout, (
            f"PowerPointFilterTest not mentioned in Maven output.\n"
            f"Maven stdout (last 2000 chars):\n{stdout[-2000:]}"
        )
