/**
 * Standalone comparison test for PowerPointFilter vs TikaTextExtractionFilter.
 * This does NOT extend AbstractUnitTest to avoid DSpace kernel initialization.
 */
package org.dspace.app.mediafilter;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;

public class FilterComparisonTest {

    public static void main(String[] args) throws Exception {
        if (args.length < 2) {
            System.err.println("Usage: FilterComparisonTest <ppt|pptx> <fixture-path>");
            System.exit(1);
        }

        String format = args[0];  // "ppt" or "pptx"
        String fixturePath = args[1];

        // Read the fixture file
        byte[] fileBytes = Files.readAllBytes(Paths.get(fixturePath));

        // Test PowerPointFilter (target)
        PowerPointFilter ppFilter = new PowerPointFilter();
        InputStream ppInput = new ByteArrayInputStream(fileBytes);
        InputStream ppOutput = ppFilter.getDestinationStream(null, ppInput, false);
        String ppText = readStream(ppOutput);

        // Test TikaTextExtractionFilter (origin baseline)
        TikaTextExtractionFilter tikaFilter = new TikaTextExtractionFilter();
        InputStream tikaInput = new ByteArrayInputStream(fileBytes);
        InputStream tikaOutput = tikaFilter.getDestinationStream(null, tikaInput, false);
        String tikaText = readStream(tikaOutput);

        // Output results in a parseable format
        System.out.println("FORMAT:" + format);
        System.out.println("POWERPOINT_LENGTH:" + (ppText != null ? ppText.length() : 0));
        System.out.println("TIKA_LENGTH:" + (tikaText != null ? tikaText.length() : 0));
        System.out.println("---POWERPOINT_OUTPUT---");
        System.out.println(ppText != null ? ppText : "(null)");
        System.out.println("---TIKA_OUTPUT---");
        System.out.println(tikaText != null ? tikaText : "(null)");
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
