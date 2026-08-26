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
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.dspace.content.Item;

/**
 * Media filter that extracts plain text from Microsoft PowerPoint presentations
 * (.ppt and .pptx) using Apache POI directly, preserving slide-by-slide
 * extraction semantics from the DSpace 6 implementation.
 *
 * <p>For .pptx files, Apache POI's XSLF API (XMLSlideShow) is used.
 * For .ppt files, the HSLF API (HSLFSlideShow) is used.
 * Text runs are collected in slide order and written to the TEXT bundle
 * as a UTF-8 plain-text stream.
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
     * <p>Attempts to open the source as PPTX (OOXML) first; if that fails,
     * falls back to PPT (OLE2 binary). Returns null for null or empty source,
     * or when no text is found, matching the TikaTextExtractionFilter convention
     * that {@code MediaFilterServiceImpl} uses to skip bitstreams.
     *
     * @param currentItem the item containing the bitstream (may be null in unit tests)
     * @param source      input stream of the PowerPoint file
     * @param verbose     if true, prints extracted text to STDOUT
     * @return UTF-8 encoded plain text as an InputStream, or null if no text found
     * @throws Exception if the source cannot be parsed as either PPT or PPTX
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

        try {
            StringBuilder text = new StringBuilder();
            boolean extractedAsPptx = tryExtractPptx(data, text);
            if (!extractedAsPptx) {
                tryExtractPpt(data, text);
            }

            String extractedText = text.toString();
            if (extractedText.isEmpty()) {
                return null;
            }

            if (verbose) {
                System.out.println("(Verbose mode) Extracted text:");
                System.out.println(extractedText);
            }

            return new ByteArrayInputStream(extractedText.getBytes(StandardCharsets.UTF_8));
        } catch (Exception e) {
            log.error("Error extracting text from PowerPoint bitstream", e);
            throw e;
        }
    }

    /**
     * Attempt to extract text from PPTX (OOXML) format using the XSLF API.
     *
     * @param data  raw bytes of the presentation file
     * @param text  buffer to append extracted text into
     * @return true if the file was successfully parsed as PPTX, false otherwise
     */
    private boolean tryExtractPptx(byte[] data, StringBuilder text) {
        try (XMLSlideShow pptx = new XMLSlideShow(new ByteArrayInputStream(data))) {
            for (XSLFSlide slide : pptx.getSlides()) {
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape textShape = (XSLFTextShape) shape;
                        for (XSLFTextParagraph para : textShape.getTextParagraphs()) {
                            for (XSLFTextRun run : para.getTextRuns()) {
                                String runText = run.getRawText();
                                if (runText != null) {
                                    text.append(runText);
                                }
                            }
                        }
                        text.append("\n");
                    }
                }
            }
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * Extract text from PPT (OLE2 binary) format using the HSLF API.
     *
     * @param data  raw bytes of the presentation file
     * @param text  buffer to append extracted text into
     * @throws Exception if the file cannot be parsed as a valid PPT document
     */
    private void tryExtractPpt(byte[] data, StringBuilder text) throws Exception {
        try (HSLFSlideShow ppt = new HSLFSlideShow(new ByteArrayInputStream(data))) {
            for (HSLFSlide slide : ppt.getSlides()) {
                for (List<HSLFTextParagraph> paraGroup : slide.getTextParagraphs()) {
                    for (HSLFTextParagraph para : paraGroup) {
                        for (HSLFTextRun run : para.getTextRuns()) {
                            String runText = run.getRawText();
                            if (runText != null) {
                                text.append(runText);
                            }
                        }
                    }
                    text.append("\n");
                }
            }
        }
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
