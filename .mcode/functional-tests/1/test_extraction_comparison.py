"""
Origin vs Target extraction comparison tests.

These tests compare the PowerPointFilter (target) text extraction output
against the expected output files (test.ppt.txt, test.pptx.txt) which represent
the origin baseline behavior.
"""

import subprocess
import os
import difflib

WORKSPACE_DIR = os.environ.get("WORKSPACE_DIR", "/l2l/workspace")
DSPACE_DIR = os.path.join(WORKSPACE_DIR, "DSpace")
MEDIAFILTER_RESOURCES = os.path.join(
    DSPACE_DIR,
    "dspace-api/src/test/resources/org/dspace/app/mediafilter"
)

def normalize_whitespace(text):
    """Normalize whitespace for comparison - collapse multiple spaces/newlines."""
    import re
    # Collapse multiple newlines to single newline
    text = re.sub(r'\n\n+', '\n\n', text)
    # Collapse multiple spaces to single space
    text = re.sub(r'  +', ' ', text)
    return text.strip()

def extract_with_java_code(ppt_path):
    """
    Extract text using inline Java code that calls PowerPointFilter directly.
    This avoids the complexity of classpath management.
    """
    java_code = f"""
import java.io.*;
import java.nio.file.*;
import java.nio.charset.StandardCharsets;
import org.dspace.app.mediafilter.PowerPointFilter;

public class Extract {{
    public static void main(String[] args) throws Exception {{
        PowerPointFilter filter = new PowerPointFilter();
        byte[] fileBytes = Files.readAllBytes(Paths.get("{ppt_path}"));
        InputStream input = new ByteArrayInputStream(fileBytes);
        InputStream output = filter.getDestinationStream(null, input, false);
        if (output == null) {{
            System.out.println("(null output)");
        }} else {{
            byte[] bytes = output.readAllBytes();
            String text = new String(bytes, StandardCharsets.UTF_8);
            System.out.print(text);
        }}
    }}
}}
"""

    # Write Java code to temp file
    java_file = "/tmp/Extract.java"
    with open(java_file, "w") as f:
        f.write(java_code)

    # Build complete classpath
    with open("/tmp/dspace-classpath.txt", "r") as f:
        maven_cp = f.read().strip()
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

    # Run
    run_result = subprocess.run(
        ["java", "-cp", f"/tmp:{classpath}", "Extract"],
        capture_output=True,
        text=True,
        timeout=30
    )

    if run_result.returncode != 0:
        raise RuntimeError(f"Execution failed: {run_result.stderr}")

    return run_result.stdout


class TestExtractionComparison:
    """Origin vs Target extraction comparison."""

    def test_ppt_extraction_output(self):
        """Extract text from test.ppt: PowerPointFilter (target) vs origin baseline."""
        ppt_path = os.path.join(MEDIAFILTER_RESOURCES, "test.ppt")
        expected_path = os.path.join(MEDIAFILTER_RESOURCES, "test.ppt.txt")

        # Get target output
        target_output = extract_with_java_code(ppt_path)

        # Get origin baseline
        with open(expected_path, "r", encoding="utf-8") as f:
            origin_output = f.read()

        # Normalize for comparison
        target_norm = normalize_whitespace(target_output)
        origin_norm = normalize_whitespace(origin_output)

        # Check if they match (after normalization)
        match = (target_norm == origin_norm)

        # Record results
        if not match:
            diff = list(difflib.unified_diff(
                origin_norm.splitlines(keepends=True),
                target_norm.splitlines(keepends=True),
                fromfile="origin (test.ppt.txt)",
                tofile="target (PowerPointFilter)",
                lineterm=""
            ))
            diff_text = "".join(diff[:50])  # First 50 lines of diff
            print(f"Mismatch detected. Diff (first 50 lines):\n{diff_text}")

        assert match, f"PPT extraction output mismatch. Origin length: {len(origin_norm)}, Target length: {len(target_norm)}"

        return {
            "match": match,
            "origin_output": origin_output,
            "target_output": target_output,
            "origin_length": len(origin_output),
            "target_length": len(target_output)
        }

    def test_pptx_extraction_output(self):
        """Extract text from test.pptx: PowerPointFilter (target) vs origin baseline."""
        pptx_path = os.path.join(MEDIAFILTER_RESOURCES, "test.pptx")
        expected_path = os.path.join(MEDIAFILTER_RESOURCES, "test.pptx.txt")

        # Get target output
        target_output = extract_with_java_code(pptx_path)

        # Get origin baseline
        with open(expected_path, "r", encoding="utf-8") as f:
            origin_output = f.read()

        # Normalize for comparison
        target_norm = normalize_whitespace(target_output)
        origin_norm = normalize_whitespace(origin_output)

        # Check if they match
        match = (target_norm == origin_norm)

        if not match:
            diff = list(difflib.unified_diff(
                origin_norm.splitlines(keepends=True),
                target_norm.splitlines(keepends=True),
                fromfile="origin (test.pptx.txt)",
                tofile="target (PowerPointFilter)",
                lineterm=""
            ))
            diff_text = "".join(diff[:50])
            print(f"Mismatch detected. Diff (first 50 lines):\n{diff_text}")

        assert match, f"PPTX extraction output mismatch. Origin length: {len(origin_norm)}, Target length: {len(target_norm)}"

        return {
            "match": match,
            "origin_output": origin_output,
            "target_output": target_output,
            "origin_length": len(origin_output),
            "target_length": len(target_output)
        }
