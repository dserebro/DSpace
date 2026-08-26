/**
 * The contents of this file are subject to the license and copyright
 * detailed in the LICENSE and NOTICE files at the root of the source
 * tree and available online at
 *
 * http://www.dspace.org/license/
 */
package org.dspace.app.mediafilter;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.extractor.ExtractorFactory;
import org.apache.poi.extractor.POITextExtractor;
import org.apache.poi.sl.extractor.SlideShowExtractor;
import org.apache.poi.xslf.extractor.XSLFExtractor;
import org.dspace.content.Item;

/**
 * Media filter that extracts plain text from Microsoft PowerPoint presentations
 * (.ppt and .pptx) using Apache POI directly, achieving strict behavioral parity
 * with the DSpace 6 PowerPointFilter (see git commit {@code ec451db2e1}).
 *
 * <p>The implementation mirrors the DSpace 6 approach: Apache POI's
 * {@link ExtractorFactory} auto-detects the presentation format (OLE2 binary
 * {@code .ppt} or OOXML {@code .pptx}), then both slide body text and speaker
 * notes are extracted with {@code setSlidesByDefault(true)} and
 * {@code setNotesByDefault(true)}.
 * Extracted text is returned as a UTF-8 plain-text stream stored in the TEXT bundle.
 *
 * <p>{@link SelfRegisterInputFormats} is implemented so that
 * {@code MediaFilterServiceImpl} can also match bitstreams by MIME type, providing
 * defence-in-depth for bitstreams whose format name was not detected at ingest.
 */
public class PowerPointFilter extends MediaFilter implements SelfRegisterInputFormats {

    private static final Logger log = LogManager.getLogger();

    @Override
    public String getFilteredName(String oldFilename) {
        return oldFilename + ".txt";
    }

    @Override
    public String getBundleName() {
        return "TEXT";
    }

    @Override
    public String getFormatString() {
        return "Text";
    }

    @Override
    public String getDescription() {
        return "Extracted text";
    }

    /**
     * Extract plain text from a PowerPoint presentation.
     *
     * <p>Returns null for null or empty source, or when no text is found,
     * matching the TikaTextExtractionFilter convention that
     * {@code MediaFilterServiceImpl} uses to skip bitstreams.
     *
     * @param currentItem the item containing the bitstream (may be null in unit tests)
     * @param source      input stream of the PowerPoint file
     * @param verbose     if true, prints extracted text to STDOUT
     * @return UTF-8 encoded plain text as an InputStream, or null if no text found
     * @throws Exception if the source cannot be parsed
     */
    @Override
    public InputStream getDestinationStream(Item currentItem, InputStream source, boolean verbose)
        throws Exception {
        if (source == null) {
            return null;
        }

        byte[] data = source.readAllBytes();
        if (data.length == 0) {
            return null;
        }

        String extractedText = null;

        try (POITextExtractor extractor = ExtractorFactory.createExtractor(new ByteArrayInputStream(data))) {
            if (extractor instanceof XSLFExtractor) {
                XSLFExtractor xslfExtractor = (XSLFExtractor) extractor;
                xslfExtractor.setSlidesByDefault(true);
                xslfExtractor.setNotesByDefault(true);
                extractedText = xslfExtractor.getText();
            } else if (extractor instanceof SlideShowExtractor) {
                @SuppressWarnings("unchecked")
                SlideShowExtractor<?, ?> slideExtractor = (SlideShowExtractor<?, ?>) extractor;
                slideExtractor.setSlidesByDefault(true);
                slideExtractor.setNotesByDefault(true);
                extractedText = slideExtractor.getText();
            }
        } catch (Exception e) {
            log.error("Error extracting text from PowerPoint bitstream", e);
            throw e;
        }

        if (extractedText == null || extractedText.isEmpty()) {
            return null;
        }

        if (verbose) {
            System.out.println("(Verbose mode) Extracted text:");
            System.out.println(extractedText);
        }

        return new ByteArrayInputStream(extractedText.getBytes(StandardCharsets.UTF_8));
    }

    @Override
    public String[] getInputMIMETypes() {
        return new String[] {
            "application/vnd.ms-powerpoint",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        };
    }

    @Override
    public String[] getInputDescriptions() {
        return null;
    }

    @Override
    public String[] getInputExtensions() {
        return new String[] {"ppt", "pptx"};
    }
}
