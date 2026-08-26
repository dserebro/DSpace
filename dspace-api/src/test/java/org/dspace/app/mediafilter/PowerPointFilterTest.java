/**
 * The contents of this file are subject to the license and copyright
 * detailed in the LICENSE and NOTICE files at the root of the source
 * tree and available online at
 *
 * http://www.dspace.org/license/
 */
package org.dspace.app.mediafilter;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

import org.apache.commons.io.IOUtils;
import org.dspace.AbstractUnitTest;
import org.junit.Before;
import org.junit.Test;

/**
 * Unit tests for {@link PowerPointFilter}, validating slide-text extraction from
 * both binary .ppt and OOXML .pptx fixtures.
 *
 * <p>Fixture files are located at
 * {@code dspace-api/src/test/resources/org/dspace/app/mediafilter/} and are
 * the same files used by {@link TikaTextExtractionFilterTest}.
 */
public class PowerPointFilterTest extends AbstractUnitTest {

    private PowerPointFilter filter;

    @Before
    public void setUp() throws Exception {
        super.init();
        filter = new PowerPointFilter();
    }

    /**
     * Verify that getFilteredName appends ".txt" to the source filename.
     */
    @Test
    public void testGetFilteredName() {
        assertTrue(filter.getFilteredName("presentation.pptx").endsWith(".txt"));
    }

    /**
     * Verify that getBundleName returns "TEXT".
     */
    @Test
    public void testGetBundleName() {
        assertTrue("TEXT".equals(filter.getBundleName()));
    }

    /**
     * Verify that getFormatString returns "Text".
     */
    @Test
    public void testGetFormatString() {
        assertTrue("Text".equals(filter.getFormatString()));
    }

    /**
     * Extract text from the binary .ppt fixture and assert the known slide content
     * "quick brown fox" is present — the same phrase verified by TikaTextExtractionFilterTest.
     */
    @Test
    public void testGetDestinationStreamWithPPT() throws Exception {
        InputStream source = getClass().getResourceAsStream("test.ppt");
        assertNotNull("test.ppt fixture must exist", source);
        InputStream result = filter.getDestinationStream(null, source, false);
        assertNotNull("PPT extraction should return a non-null stream", result);
        String text = readAll(result);
        assertTrue("Known content 'quick brown fox' not found in .ppt output", text.contains("quick brown fox"));
    }

    /**
     * Extract text from the OOXML .pptx fixture and assert the known slide content
     * "quick brown fox" is present — the same phrase verified by TikaTextExtractionFilterTest.
     */
    @Test
    public void testGetDestinationStreamWithPPTX() throws Exception {
        InputStream source = getClass().getResourceAsStream("test.pptx");
        assertNotNull("test.pptx fixture must exist", source);
        InputStream result = filter.getDestinationStream(null, source, false);
        assertNotNull("PPTX extraction should return a non-null stream", result);
        String text = readAll(result);
        assertTrue("Known content 'quick brown fox' not found in .pptx output", text.contains("quick brown fox"));
    }

    /**
     * Verify that verbose mode prints extracted text to STDOUT without throwing.
     * The returned stream must still be non-null and contain the expected content.
     */
    @Test
    public void testVerboseMode() throws Exception {
        InputStream source = getClass().getResourceAsStream("test.pptx");
        assertNotNull("test.pptx fixture must exist", source);
        InputStream result = filter.getDestinationStream(null, source, true);
        assertNotNull("Verbose mode should still return a non-null stream", result);
        assertTrue(readAll(result).contains("quick brown fox"));
    }

    /**
     * Verify that a null source stream returns null, matching the TikaTextExtractionFilter
     * convention that MediaFilterServiceImpl uses to skip the bitstream.
     */
    @Test
    public void testNullSourceReturnsNull() throws Exception {
        InputStream result = filter.getDestinationStream(null, null, false);
        assertNull("null source should return null", result);
    }

    /**
     * Verify that an empty (zero-byte) source stream returns null, matching the spec
     * requirement to handle both null and empty sources gracefully.
     */
    @Test
    public void testEmptySourceReturnsNull() throws Exception {
        InputStream result = filter.getDestinationStream(null, new ByteArrayInputStream(new byte[0]), false);
        assertNull("empty source should return null", result);
    }

    /**
     * Verify that SelfRegisterInputFormats returns the expected PPT and PPTX MIME types.
     */
    @Test
    public void testGetInputMIMETypes() {
        String[] mimeTypes = filter.getInputMIMETypes();
        assertNotNull(mimeTypes);
        assertTrue("Should include application/vnd.ms-powerpoint",
                containsValue(mimeTypes, "application/vnd.ms-powerpoint"));
        assertTrue("Should include application/vnd.openxmlformats-officedocument.presentationml.presentation",
                containsValue(mimeTypes,
                        "application/vnd.openxmlformats-officedocument.presentationml.presentation"));
    }

    private static boolean containsValue(String[] array, String value) {
        for (String s : array) {
            if (value.equals(s)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Character-for-character output-stability test for .ppt extraction.
     *
     * <p>Compares the filter output against the committed expected-output fixture
     * {@code test.ppt.txt}. The fixture was generated by running this implementation
     * (direct HSLF API, as mandated by the milestone design decisions) against the binary
     * PPT test file; the DSpace 6 source is not present in this workspace.
     * Any future deviation in slide-text ordering, whitespace, or content will fail
     * this test and require deliberate fixture regeneration.
     */
    @Test
    public void testParityWithExpectedOutputPPT() throws Exception {
        InputStream source = getClass().getResourceAsStream("test.ppt");
        assertNotNull("test.ppt fixture must exist", source);
        InputStream result = filter.getDestinationStream(null, source, false);
        assertNotNull("PPT extraction must return a non-null stream", result);
        String actual = readAll(result);

        InputStream expected = getClass().getResourceAsStream("test.ppt.txt");
        assertNotNull("test.ppt.txt expected-output fixture must exist", expected);
        String expectedText = readAll(expected);

        assertEquals("PPT extraction output must match expected fixture character-for-character",
                expectedText, actual);
    }

    /**
     * Character-for-character output-stability test for .pptx extraction.
     *
     * <p>Compares the filter output against the committed expected-output fixture
     * {@code test.pptx.txt}. The fixture was generated by running this implementation
     * (direct XSLF API, as mandated by the milestone design decisions) against the OOXML
     * PPTX test file; the DSpace 6 source is not present in this workspace.
     * Any future deviation in slide-text ordering, whitespace, or content will fail
     * this test and require deliberate fixture regeneration.
     */
    @Test
    public void testParityWithExpectedOutputPPTX() throws Exception {
        InputStream source = getClass().getResourceAsStream("test.pptx");
        assertNotNull("test.pptx fixture must exist", source);
        InputStream result = filter.getDestinationStream(null, source, false);
        assertNotNull("PPTX extraction must return a non-null stream", result);
        String actual = readAll(result);

        InputStream expected = getClass().getResourceAsStream("test.pptx.txt");
        assertNotNull("test.pptx.txt expected-output fixture must exist", expected);
        String expectedText = readAll(expected);

        assertEquals("PPTX extraction output must match expected fixture character-for-character",
                expectedText, actual);
    }

    private static String readAll(InputStream stream) throws IOException {
        return IOUtils.toString(stream, StandardCharsets.UTF_8);
    }
}
