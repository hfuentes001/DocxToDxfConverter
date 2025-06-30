package org.teknei;

import org.apache.poi.xwpf.usermodel.*;
import javax.xml.stream.*;
import java.io.*;
import java.math.BigInteger;
import java.nio.file.*;
import java.util.*;
import java.util.Base64;

public class DocxToDxfConverter {
    // Contadores por lista (numID + nivel)
    private static final Map<String, Integer> numCounters = new HashMap<>();

    private static class RunData {
        String text;
        String font;
        int size;
        boolean bold;
        boolean italic;
        boolean underline;
        boolean strikethrough;

        // Para párrafos: prefix contendrá "• " o "1. ", "2. ", etc.
        String prefix;
        boolean isParagraph;
        ParagraphAlignment alignment;

        // Run normal
        RunData(String text, String font, int size,
                boolean bold, boolean italic,
                boolean underline, boolean strikethrough) {
            this.text = text;
            this.font = font;
            this.size = size;
            this.bold = bold;
            this.italic = italic;
            this.underline = underline;
            this.strikethrough = strikethrough;
            this.isParagraph = false;
        }

        // Párrafo de lista
        RunData(ParagraphAlignment alignment, String prefix) {
            this.isParagraph = true;
            this.alignment = alignment;
            this.prefix = prefix;
        }
    }

    public static void main(String[] args) throws Exception {
        if (args.length == 0) {
            System.err.println("Uso: java DocxToDxfConverter <archivo.docx>");
            return;
        }
        Path docxPath = Paths.get(args[0]);
        Path outputDir = Paths.get("salida");
        Files.createDirectories(outputDir);

        try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(docxPath))) {
            List<RunData> buffer = new ArrayList<>();
            int counter = 1;

            for (IBodyElement element : doc.getBodyElements()) {
                if (element instanceof XWPFTable) {
                    if (!buffer.isEmpty()) {
                        flushBuffer(buffer, outputDir.resolve(counter++ + ".dxf"));
                        buffer.clear();
                    }
                    writeTable(outputDir.resolve(counter++ + ".dxf"), (XWPFTable) element);
                }
                else if (element instanceof XWPFParagraph) {
                    XWPFParagraph para = (XWPFParagraph) element;
                    String prefix = computeListPrefix(doc, para);
                    buffer.add(new RunData(para.getAlignment(), prefix));

                    for (XWPFRun run : para.getRuns()) {
                        for (XWPFPicture pic : run.getEmbeddedPictures()) {
                            int picType = pic.getPictureData().getPictureType();
                            Integer rot = pic.getCTPicture()
                                    .getSpPr().getXfrm().getRot();
                            if (picType == Document.PICTURE_TYPE_JPEG && (rot == null || rot == 0)) {
                                if (!buffer.isEmpty()) {
                                    flushBuffer(buffer, outputDir.resolve(counter++ + ".dxf"));
                                    buffer.clear();
                                }
                                writeImage(outputDir.resolve(counter++ + ".dxf"), pic);
                            }
                        }
                        String text = run.getText(0);
                        if (text != null && !text.isEmpty()) {
                            String fontFamily = Optional.ofNullable(run.getFontFamily())
                                    .orElse("Arial");
                            int fontSize = run.getFontSize() > 0 ? run.getFontSize() : 10;
                            boolean under = run.getUnderline() != UnderlinePatterns.NONE;
                            boolean strike = run.isStrike();
                            buffer.add(new RunData(text, fontFamily, fontSize,
                                    run.isBold(), run.isItalic(), under, strike));
                        }
                    }
                }
            }
            if (!buffer.isEmpty()) {
                flushBuffer(buffer, outputDir.resolve(counter++ + ".dxf"));
            }
        }
    }

    /**
     * Devuelve el texto de prefijo según el tipo de lista:
     *  - bullet  → "• "
     *  - decimal → "1. ", "2. ", etc.
     */
    private static String computeListPrefix(XWPFDocument doc, XWPFParagraph para) {
        BigInteger numID = para.getNumID();
        if (numID == null) return null;

        String fmt = para.getNumFmt();  // "bullet", "decimal", etc.
        BigInteger ilvlBI = para.getNumIlvl();
        int ilvl = ilvlBI != null ? ilvlBI.intValue() : 0;
        String key = numID.toString() + "-" + ilvl;
        int count = numCounters.getOrDefault(key, 0) + 1;
        numCounters.put(key, count);

        if ("bullet".equalsIgnoreCase(fmt)) {
            return "\u2022 ";
        } else {
            return count + ". ";
        }
    }

    private static void flushBuffer(List<RunData> buffer, Path outFile) throws Exception {
        try (OutputStream os = Files.newOutputStream(outFile)) {
            writeRunsDXF(buffer, os);
        }
    }

    private static void writeRunsDXF(List<RunData> runs, OutputStream os) throws Exception {
        XMLStreamWriter w = XMLOutputFactory.newInstance()
                .createXMLStreamWriter(os, "UTF-8");
        w.writeStartDocument("UTF-8", "1.0");
        w.writeDTD("<!DOCTYPE dlg:dxf-text SYSTEM \"ExstreamObjectAndContent.dtd\">");
        w.writeStartElement("dlg:dxf-text");
        w.writeDefaultNamespace("http://www.exstream.com/2003/XSL/Dialogue");
        w.writeNamespace("fo", "http://www.w3.org/1999/XSL/Format");
        w.writeStartElement("fo:flow");
        w.writeAttribute("height", "755.00pt");
        w.writeAttribute("width", "570.00pt");
        w.writeAttribute("margin-top", "0.00pt");
        w.writeAttribute("margin-bottom", "0.00pt");
        w.writeAttribute("margin-left", "0.00pt");
        w.writeAttribute("margin-right", "0.00pt");

        boolean blockOpen = false;
        for (RunData rd : runs) {
            if (rd.isParagraph) {
                if (blockOpen) w.writeEndElement();
                w.writeStartElement("fo:block");
                w.writeAttribute("line-spacing", "at-least");
                String align = (rd.alignment == ParagraphAlignment.CENTER) ? "center"
                        : (rd.alignment == ParagraphAlignment.RIGHT) ? "right"
                        : (rd.alignment == ParagraphAlignment.BOTH
                        || rd.alignment == ParagraphAlignment.DISTRIBUTE) ? "justify"
                        : "left";
                w.writeAttribute("text-align", align);

                if (rd.prefix != null) {
                    w.writeStartElement("fo:inline");
                    w.writeCharacters(rd.prefix);
                    w.writeEndElement();
                }

                blockOpen = true;
            }
            else {
                w.writeStartElement("fo:inline");
                w.writeAttribute("color", "rgb(0,0,0)");
                w.writeAttribute("font-family", rd.font);
                w.writeAttribute("font-size", rd.size + "pt");
                if (rd.bold)          w.writeAttribute("font-weight", "bold");
                if (rd.italic)        w.writeAttribute("font-style", "italic");
                if (rd.underline)     w.writeAttribute("text-decoration", "underline");
                if (rd.strikethrough) w.writeAttribute("text-decoration", "line-through");
                w.writeCharacters(rd.text);
                w.writeEndElement();
            }
        }
        if (blockOpen) w.writeEndElement();
        w.writeEndElement(); // fo:flow
        w.writeEndElement(); // dlg:dxf-text
        w.writeEndDocument();
        w.close();
    }

    private static void writeImage(Path path, XWPFPicture pic) throws Exception {
        long cx = pic.getCTPicture().getSpPr().getXfrm().getExt().getCx();
        long cy = pic.getCTPicture().getSpPr().getXfrm().getExt().getCy();
        double widthPt  = cx / 12700.0;
        double heightPt = cy / 12700.0;

        try (OutputStream os = Files.newOutputStream(path)) {
            XMLStreamWriter w = XMLOutputFactory.newInstance()
                    .createXMLStreamWriter(os, "UTF-8");
            w.writeStartDocument("UTF-8", "1.0");
            w.writeDTD("<!DOCTYPE dlg:library-component SYSTEM \"ExstreamObjectAndContent.dtd\">");
            w.writeStartElement("dlg:library-component");
            w.writeDefaultNamespace("http://www.exstream.com/2003/XSL/Dialogue");
            w.writeStartElement("dlg:object");
            w.writeStartElement("dlg:image");
            w.writeStartElement("dlg:rect");
            w.writeAttribute("bottom", String.format("%.2fpt", heightPt));
            w.writeAttribute("left",   "0.00pt");
            w.writeAttribute("right",  String.format("%.2fpt", widthPt));
            w.writeAttribute("top",    "0.00pt");
            w.writeEndElement();
            w.writeStartElement("dlg:bitmap");
            w.writeStartElement("dlg:binary");
            w.writeAttribute("encoding", "base64");
            byte[] data = pic.getPictureData().getData();
            w.writeCharacters(Base64.getEncoder().encodeToString(data));
            w.writeEndElement();
            w.writeEndElement();
            w.writeEndElement();
            w.writeEndElement();
            w.writeEndDocument();
            w.close();
        }
    }

    private static void writeTable(Path path, XWPFTable table) throws Exception {
        List<XWPFTableRow> rows = table.getRows();
        if (rows.isEmpty()) return;

        // Columnas
        XWPFTableRow first = rows.get(0);
        List<Double> colWidths = new ArrayList<>();
        double tableWidth = 0;
        for (XWPFTableCell cell : first.getTableCells()) {
            double pts = cell.getWidth() / 20.0;
            colWidths.add(pts);
            tableWidth += pts;
        }

        // Altura total
        double tableHeight = rows.stream()
                .mapToDouble(r -> r.getHeight() > 0 ? r.getHeight()/20.0 : 12.0)
                .sum();

        try (OutputStream os = Files.newOutputStream(path)) {
            XMLStreamWriter w = XMLOutputFactory.newInstance()
                    .createXMLStreamWriter(os, "UTF-8");
            w.writeStartDocument("UTF-8", "1.0");
            w.writeDTD("<!DOCTYPE dlg:library-component SYSTEM \"ExstreamObjectAndContent.dtd\">");
            w.writeStartElement("dlg:library-component");
            w.writeDefaultNamespace("http://www.exstream.com/2003/XSL/Dialogue");
            w.writeStartElement("dlg:object");
            w.writeStartElement("dlg:table");

            // Rect
            w.writeStartElement("dlg:rect");
            w.writeAttribute("bottom", String.format("%.2fpt", tableHeight));
            w.writeAttribute("left",   "0.00pt");
            w.writeAttribute("right",  String.format("%.2fpt", tableWidth));
            w.writeAttribute("top",    "0.00pt");
            w.writeEndElement();

            // Definir columnas
            for (int i = 0; i < colWidths.size(); i++) {
                w.writeStartElement("fo:table-column");
                w.writeAttribute("column-width", String.format("%.2fpt", colWidths.get(i)));
                w.writeAttribute("column-number", String.valueOf(i+1));
                w.writeEndElement();
            }

            // Filas y celdas
            for (XWPFTableRow row : rows) {
                w.writeStartElement("fo:table-row");
                w.writeAttribute("border-top-color",   "rgb(0,0,0)");
                w.writeAttribute("border-top-style",   "solid");
                w.writeAttribute("border-top-width",   "0.24pt");
                w.writeAttribute("border-bottom-color","rgb(0,0,0)");
                w.writeAttribute("border-bottom-style","solid");
                w.writeAttribute("border-bottom-width","0.24pt");

                List<XWPFTableCell> cells = row.getTableCells();
                for (int ci = 0; ci < cells.size(); ci++) {
                    XWPFTableCell cell = cells.get(ci);
                    double cw = colWidths.get(ci);

                    w.writeStartElement("fo:table-cell");
                    w.writeAttribute("display-align","before");
                    w.writeAttribute("width", String.format("%.2fpt", cw));
                    w.writeAttribute("margin-left","5.52pt");
                    w.writeAttribute("margin-right","5.52pt");
                    w.writeAttribute("column-number", String.valueOf(ci+1));
                    w.writeAttribute("border-left-color","rgb(0,0,0)");
                    w.writeAttribute("border-left-style","solid");
                    w.writeAttribute("border-left-width","0.24pt");
                    if (ci == cells.size()-1) {
                        w.writeAttribute("border-right-color","rgb(0,0,0)");
                        w.writeAttribute("border-right-style","solid");
                        w.writeAttribute("border-right-width","0.24pt");
                    }

                    // Contenido de la celda
                    for (XWPFParagraph para : cell.getParagraphs()) {
                        w.writeStartElement("fo:block");
                        w.writeAttribute("tab-ruler","-1");
                        ParagraphAlignment pa = para.getAlignment();
                        String align = (pa == ParagraphAlignment.CENTER) ? "center"
                                : (pa == ParagraphAlignment.RIGHT) ? "right"
                                : (pa == ParagraphAlignment.BOTH
                                || pa == ParagraphAlignment.DISTRIBUTE) ? "justify"
                                : "left";
                        w.writeAttribute("text-align", align);

                        String prefix = computeListPrefix(null, para);
                        if (prefix != null) {
                            w.writeStartElement("fo:inline");
                            w.writeCharacters(prefix);
                            w.writeEndElement();
                        }

                        for (XWPFRun run : para.getRuns()) {
                            String txt = run.getText(0);
                            if (txt == null || txt.isEmpty()) continue;

                            w.writeStartElement("fo:inline");
                            String hex = run.getColor();
                            w.writeAttribute("color",
                                    hex!=null&&hex.length()==6
                                            ? "rgb(" + hexToRgb(hex) + ")"
                                            : "rgb(0,0,0)");
                            String fam = Optional.ofNullable(run.getFontFamily())
                                    .orElse("Arial");
                            w.writeAttribute("font-family", fam);
                            int sz = run.getFontSize()>0?run.getFontSize():10;
                            w.writeAttribute("font-size", sz+"pt");
                            if (run.isBold())   w.writeAttribute("font-weight","bold");
                            if (run.isItalic()) w.writeAttribute("font-style","italic");
                            if (run.getUnderline()!=UnderlinePatterns.NONE) {
                                w.writeAttribute("text-decoration","underline");
                            }
                            if (run.isStrike()) w.writeAttribute("text-decoration","line-through");
                            w.writeCharacters(txt);
                            w.writeEndElement();
                        }
                        w.writeEndElement(); // fo:block
                    }

                    w.writeEndElement(); // fo:table-cell
                }
                w.writeEndElement(); // fo:table-row
            }

            w.writeEndElement(); // dlg:table
            w.writeEndElement(); // dlg:object
            w.writeEndElement(); // dlg:library-component
            w.writeEndDocument();
            w.close();
        }
    }

    private static String hexToRgb(String hex) {
        int r = Integer.parseInt(hex.substring(0,2),16);
        int g = Integer.parseInt(hex.substring(2,4),16);
        int b = Integer.parseInt(hex.substring(4,6),16);
        return r + "," + g + "," + b;
    }
}
