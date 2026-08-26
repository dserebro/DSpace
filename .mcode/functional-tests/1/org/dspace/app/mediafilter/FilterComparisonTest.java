/**
 * Standalone comparison test for PowerPointFilter vs expected-output baseline.
 * This does NOT extend AbstractUnitTest to avoid DSpace kernel initialization.
 * The origin baseline is read from the committed fixture file rather than invoking
 * TikaTextExtractionFilter (which requires a running DSpace kernel).
 */
package org.dspace.app.mediafilter;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;

public class FilterComparisonTest {

    public static void main(String[] args) throws Exception {
        if (args.length < 3) {
            System.err.println("Usage: FilterComparisonTest <ppt|pptx> <fixture-path> <expected-output-path>");
            System.exit(1);
        }

        String format = args[0];       // "ppt" or "pptx"
        String fixturePath = args[1];  // path to test.ppt or test.pptx
        String expectedPath = args[2]; // path to test.ppt.txt or test.pptx.txt

        // Read the fixture file
        byte[] fileBytes = Files.readAllBytes(Paths.get(fixturePath));

        // Test PowerPointFilter (target)
        PowerPointFilter ppFilter = new PowerPointFilter();
        InputStream ppInput = new ByteArrayInputStream(fileBytes);
        InputStream ppOutput = ppFilter.getDestinationStream(null, ppInput, false);
        String ppText = readStream(ppOutput);

        // Read origin baseline from committed fixture file
        byte[] expectedBytes = Files.readAllBytes(Paths.get(expectedPath));
        String expectedText = new String(expectedBytes, StandardCharsets.UTF_8);

        // Output results in a parseable format
        System.out.println("FORMAT:" + format);
        System.out.println("POWERPOINT_LENGTH:" + (ppText != null ? ppText.length() : 0));
        System.out.println("EXPECTED_LENGTH:" + expectedText.length());
        System.out.println("---POWERPOINT_OUTPUT---");
        System.out.println(ppText != null ? ppText : "(null)");
        System.out.println("---EXPECTED_OUTPUT---");
        System.out.println(expectedText);
        System.out.println("---END---");
    }

    private static String readStream(InputStream is) throws Exception {
        if (is == null) {
            return null;
        }
        byte[] bytes = is.readAllBytes();
        return new String(bytes, StandardCharsets.UTF_8);
    }
}
