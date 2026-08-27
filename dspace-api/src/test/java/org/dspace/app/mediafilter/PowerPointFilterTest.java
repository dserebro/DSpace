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
public class PowerPointFilterTest {

    private PowerPointFilter filter;

    @Before
    public void setUp() {
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
        assertExtractionContainsKnownContent("ppt");
    }

    /**
     * Extract text from the OOXML .pptx fixture and assert the known slide content
     * "quick brown fox" is present — the same phrase verified by TikaTextExtractionFilterTest.
     */
    @Test
    public void testGetDestinationStreamWithPPTX() throws Exception {
        assertExtractionContainsKnownContent("pptx");
    }

    private void assertExtractionContainsKnownContent(String ext) throws Exception {
        InputStream source = getClass().getResourceAsStream("test." + ext);
        assertNotNull("test." + ext + " fixture must exist", source);
        InputStream result = filter.getDestinationStream(null, source, false);
        assertNotNull(ext.toUpperCase() + " extraction should return a non-null stream", result);
        String text = readAll(result);
        assertTrue("Known content 'quick brown fox' not found in ." + ext + " output",
                text.contains("quick brown fox"));
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
     * Verify that corrupt (non-PowerPoint) bytes cause an exception to be thrown,
     * covering the catch block in getDestinationStream() that logs and re-throws.
     * A regression that swallows the exception would fail this test.
     */
    @Test(expected = Exception.class)
    public void testCorruptBytesThrowsException() throws Exception {
        byte[] corrupt = "THIS IS NOT A VALID POWERPOINT FILE".getBytes(StandardCharsets.UTF_8);
        filter.getDestinationStream(null, new ByteArrayInputStream(corrupt), false);
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

    /**
     * Verify that SelfRegisterInputFormats returns the expected PPT and PPTX file extensions.
     * MILESTONE.md Design Decision 2 mandates {@code "ppt"} and {@code "pptx"}.
     */
    @Test
    public void testGetInputExtensions() {
        String[] extensions = filter.getInputExtensions();
        assertNotNull(extensions);
        assertTrue("Should include 'ppt' extension", containsValue(extensions, "ppt"));
        assertTrue("Should include 'pptx' extension", containsValue(extensions, "pptx"));
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
     * Character-for-character behavioral parity test for .ppt extraction.
     *
     * <p>Compares the filter output against the committed expected-output fixture
     * {@code test.ppt.txt}, which was produced by this implementation (direct HSLF API,
     * as mandated by the milestone design decisions). Any deviation in slide-text ordering,
     * whitespace, or content signals a regression in behavioral parity and requires
     * deliberate fixture regeneration.
     */
    @Test
    public void testParityWithExpectedOutputPPT() throws Exception {
        assertParityWithExpectedOutput("ppt");
    }

    /**
     * Character-for-character behavioral parity test for .pptx extraction.
     *
     * <p>Compares the filter output against the committed expected-output fixture
     * {@code test.pptx.txt}, which was produced by this implementation (direct XSLF API,
     * as mandated by the milestone design decisions). Any deviation in slide-text ordering,
     * whitespace, or content signals a regression in behavioral parity and requires
     * deliberate fixture regeneration.
     */
    @Test
    public void testParityWithExpectedOutputPPTX() throws Exception {
        assertParityWithExpectedOutput("pptx");
    }

    private void assertParityWithExpectedOutput(String ext) throws Exception {
        InputStream source = getClass().getResourceAsStream("test." + ext);
        assertNotNull("test." + ext + " fixture must exist", source);
        InputStream result = filter.getDestinationStream(null, source, false);
        assertNotNull(ext.toUpperCase() + " extraction must return a non-null stream", result);
        String actual = readAll(result);

        InputStream expected = getClass().getResourceAsStream("test." + ext + ".txt");
        assertNotNull("test." + ext + ".txt expected-output fixture must exist", expected);
        String expectedText = readAll(expected);

        assertEquals(ext.toUpperCase() + " extraction output must match expected fixture character-for-character",
                expectedText, actual);
    }

    private static String readAll(InputStream stream) throws IOException {
        return IOUtils.toString(stream, StandardCharsets.UTF_8);
    }
}
