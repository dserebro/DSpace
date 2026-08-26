"""
Origin vs Target extraction comparison tests.

These tests compare the PowerPointFilter (target) text extraction output
against the expected output files (test.ppt.txt, test.pptx.txt) which represent
the origin baseline behavior.
"""

import subprocess
import os
import difflib
import re

WORKSPACE_DIR = os.environ.get("WORKSPACE_DIR", "/l2l/workspace")
DSPACE_DIR = os.path.join(WORKSPACE_DIR, "DSpace")
MEDIAFILTER_RESOURCES = os.path.join(
    DSPACE_DIR,
    "dspace-api/src/test/resources/org/dspace/app/mediafilter"
)
JAVA_HOME = os.environ.get("JAVA_HOME", "/usr/lib/jvm/java-21-openjdk-amd64")
MAVEN_HOME = os.environ.get("MAVEN_HOME", "/usr/share/maven")
MVN = os.path.join(MAVEN_HOME, "bin", "mvn")


def normalize_whitespace(text):
    """Normalize whitespace for comparison - collapse multiple spaces/newlines."""
    # Collapse multiple newlines to single newline
    text = re.sub(r'\n\n+', '\n\n', text)
    # Collapse multiple spaces to single space
    text = re.sub(r'  +', ' ', text)
    return text.strip()


def _build_classpath():
    """Generate the Maven dependency classpath, writing it to /tmp/dspace-classpath.txt."""
    result = subprocess.run(
        [MVN, "dependency:build-classpath", "-Dmdep.outputFile=/tmp/dspace-classpath.txt",
         "-pl", "dspace-api", "-q"],
        cwd=DSPACE_DIR,
        capture_output=True,
        text=True,
        timeout=120
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"Failed to build Maven classpath: {result.stderr}"
        )
    with open("/tmp/dspace-classpath.txt", "r") as f:
        return f.read().strip()


def extract_with_java_code(ppt_path):
    """
    Extract text using inline Java code that calls PowerPointFilter directly.
    The file path is passed via command-line argument to avoid injection risks.
    """
    java_code = """
import java.io.*;
import java.nio.file.*;
import java.nio.charset.StandardCharsets;
import org.dspace.app.mediafilter.PowerPointFilter;

public class Extract {
    public static void main(String[] args) throws Exception {
        PowerPointFilter filter = new PowerPointFilter();
        byte[] fileBytes = Files.readAllBytes(Paths.get(args[0]));
        InputStream input = new ByteArrayInputStream(fileBytes);
        InputStream output = filter.getDestinationStream(null, input, false);
        if (output == null) {
            System.out.println("(null output)");
        } else {
            byte[] bytes = output.readAllBytes();
            String text = new String(bytes, StandardCharsets.UTF_8);
            System.out.print(text);
        }
    }
}
"""

    # Write Java code to temp file
    java_file = "/tmp/Extract.java"
    with open(java_file, "w") as f:
        f.write(java_code)

    # Build complete classpath (generates /tmp/dspace-classpath.txt if needed)
    maven_cp = _build_classpath()
    classpath = f"{DSPACE_DIR}/dspace-api/target/classes:{maven_cp}"

    # Compile
    compile_result = subprocess.run(
        ["javac", "-cp", classpath, java_file],
        capture_output=True,
        text=True,
        timeout=60
    )

    if compile_result.returncode != 0:
        raise RuntimeError(f"Compilation failed: {compile_result.stderr}")

    # Run — path passed as argument, not embedded in source
    run_result = subprocess.run(
        ["java", "-cp", f"/tmp:{classpath}", "Extract", ppt_path],
        capture_output=True,
        text=True,
        timeout=30
    )

    if run_result.returncode != 0:
        raise RuntimeError(f"Execution failed: {run_result.stderr}")

    return run_result.stdout


class TestExtractionComparison:
    """Origin vs Target extraction comparison."""

    def _run_extraction_comparison(self, ext):
        """Shared helper: extract and compare PowerPointFilter output vs baseline fixture."""
        ppt_path = os.path.join(MEDIAFILTER_RESOURCES, f"test.{ext}")
        expected_path = os.path.join(MEDIAFILTER_RESOURCES, f"test.{ext}.txt")

        target_output = extract_with_java_code(ppt_path)

        with open(expected_path, "r", encoding="utf-8") as f:
            origin_output = f.read()

        target_norm = normalize_whitespace(target_output)
        origin_norm = normalize_whitespace(origin_output)

        match = (target_norm == origin_norm)

        if not match:
            diff = list(difflib.unified_diff(
                origin_norm.splitlines(keepends=True),
                target_norm.splitlines(keepends=True),
                fromfile=f"origin (test.{ext}.txt)",
                tofile=f"target (PowerPointFilter)",
                lineterm=""
            ))
            diff_text = "".join(diff[:50])
            print(f"Mismatch detected. Diff (first 50 lines):\n{diff_text}")

        assert match, (
            f"{ext.upper()} extraction output mismatch. "
            f"Origin length: {len(origin_norm)}, Target length: {len(target_norm)}"
        )

        return {
            "match": match,
            "origin_output": origin_output,
            "target_output": target_output,
            "origin_length": len(origin_output),
            "target_length": len(target_output)
        }

    def test_ppt_extraction_output(self):
        """Extract text from test.ppt: PowerPointFilter (target) vs origin baseline."""
        return self._run_extraction_comparison("ppt")

    def test_pptx_extraction_output(self):
        """Extract text from test.pptx: PowerPointFilter (target) vs origin baseline."""
        return self._run_extraction_comparison("pptx")
