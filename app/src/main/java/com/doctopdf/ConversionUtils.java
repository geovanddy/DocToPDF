package com.doctopdf;

import android.content.Context;
import android.net.Uri;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.layout.properties.UnitValue;
import com.itextpdf.layout.properties.TextAlignment;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Locale;

public class ConversionUtils {

    public static File convertToPdf(Context context, Uri uri, String fileName) throws Exception {
        String ext = FileUtils.getExtension(fileName);
        String baseName = FileUtils.removeExtension(fileName);

        // Directorio de salida en caché de la app
        File outputDir = new File(context.getCacheDir(), "pdfs");
        outputDir.mkdirs();
        File outputFile = new File(outputDir, baseName + ".pdf");

        switch (ext) {
            case "docx":
                convertDocxToPdf(context, uri, outputFile);
                break;
            case "doc":
                convertDocToPdf(context, uri, outputFile);
                break;
            case "xlsx":
                convertXlsxToPdf(context, uri, outputFile, false);
                break;
            case "xls":
                convertXlsxToPdf(context, uri, outputFile, true);
                break;
            case "pptx":
                convertPptxToPdf(context, uri, outputFile);
                break;
            case "ppt":
                convertPptToPdf(context, uri, outputFile);
                break;
            default:
                throw new Exception("Formato no soportado: " + ext);
        }

        return outputFile;
    }

    // ─── WORD DOCX → PDF ───────────────────────────────────────────────────

    private static void convertDocxToPdf(Context context, Uri uri, File output) throws Exception {
        try (InputStream is = context.getContentResolver().openInputStream(uri);
             XWPFDocument doc = new XWPFDocument(is);
             PdfWriter writer = new PdfWriter(output);
             PdfDocument pdfDoc = new PdfDocument(writer);
             Document pdfDocument = new Document(pdfDoc)) {

            // Recorrer párrafos
            for (XWPFParagraph para : doc.getParagraphs()) {
                String text = para.getText();
                if (text == null || text.trim().isEmpty()) {
                    pdfDocument.add(new Paragraph(" "));
                    continue;
                }

                Paragraph pdfPara = new Paragraph(text);

                // Detectar estilos de título
                String style = para.getStyle();
                if (style != null && style.toLowerCase().contains("heading")) {
                    pdfPara.setFontSize(16).setBold().setMarginTop(10).setMarginBottom(4);
                } else {
                    pdfPara.setFontSize(11).setMarginBottom(2);
                }

                // Alineación
                switch (para.getAlignment()) {
                    case CENTER:
                        pdfPara.setTextAlignment(TextAlignment.CENTER);
                        break;
                    case RIGHT:
                        pdfPara.setTextAlignment(TextAlignment.RIGHT);
                        break;
                    case BOTH:
                        pdfPara.setTextAlignment(TextAlignment.JUSTIFIED);
                        break;
                    default:
                        pdfPara.setTextAlignment(TextAlignment.LEFT);
                }

                pdfDocument.add(pdfPara);
            }

            // Tablas Word
            for (XWPFTable wTable : doc.getTables()) {
                int numCols = wTable.getRow(0) != null ? wTable.getRow(0).getTableCells().size() : 1;
                Table pdfTable = new Table(UnitValue.createPercentArray(numCols)).useAllAvailableWidth();

                for (XWPFTableRow row : wTable.getRows()) {
                    for (org.apache.poi.xwpf.usermodel.XWPFTableCell wCell : row.getTableCells()) {
                        String cellText = wCell.getText();
                        Cell pdfCell = new Cell().add(new Paragraph(cellText != null ? cellText : "").setFontSize(10));
                        pdfTable.addCell(pdfCell);
                    }
                }
                pdfDocument.add(pdfTable);
                pdfDocument.add(new Paragraph(" "));
            }
        }
    }

    private static void convertDocToPdf(Context context, Uri uri, File output) throws Exception {
        // Para .doc antiguo, extraemos texto plano
        try (InputStream is = context.getContentResolver().openInputStream(uri)) {
            org.apache.poi.hwpf.HWPFDocument doc = new org.apache.poi.hwpf.HWPFDocument(is);
            org.apache.poi.hwpf.extractor.WordExtractor extractor =
                new org.apache.poi.hwpf.extractor.WordExtractor(doc);
            String text = extractor.getText();

            try (PdfWriter writer = new PdfWriter(output);
                 PdfDocument pdfDoc = new PdfDocument(writer);
                 Document pdfDocument = new Document(pdfDoc)) {
                for (String line : text.split("\n")) {
                    pdfDocument.add(new Paragraph(line.trim()).setFontSize(11).setMarginBottom(2));
                }
            }
            extractor.close();
        }
    }

    // ─── EXCEL XLSX/XLS → PDF ──────────────────────────────────────────────

    private static void convertXlsxToPdf(Context context, Uri uri, File output, boolean isXls) throws Exception {
        try (InputStream is = context.getContentResolver().openInputStream(uri)) {
            Workbook workbook = isXls ? new HSSFWorkbook(is) : new XSSFWorkbook(is);

            try (PdfWriter writer = new PdfWriter(output);
                 PdfDocument pdfDoc = new PdfDocument(writer);
                 Document pdfDocument = new Document(pdfDoc)) {

                for (int si = 0; si < workbook.getNumberOfSheets(); si++) {
                    Sheet sheet = workbook.getSheetAt(si);

                    // Título de hoja
                    pdfDocument.add(new Paragraph("📊 " + sheet.getSheetName())
                        .setBold().setFontSize(14).setMarginTop(si > 0 ? 20 : 0).setMarginBottom(8)
                        .setFontColor(ColorConstants.DARK_GRAY));

                    // Calcular columnas
                    int maxCols = 0;
                    for (Row row : sheet) {
                        if (row.getLastCellNum() > maxCols) maxCols = row.getLastCellNum();
                    }
                    if (maxCols == 0) continue;

                    // Limitar a 10 columnas para que quepa en la página
                    int displayCols = Math.min(maxCols, 10);
                    Table table = new Table(UnitValue.createPercentArray(displayCols)).useAllAvailableWidth();

                    boolean firstRow = true;
                    int rowCount = 0;
                    for (Row row : sheet) {
                        if (rowCount > 500) break; // Límite razonable
                        for (int ci = 0; ci < displayCols; ci++) {
                            org.apache.poi.ss.usermodel.Cell cell = row.getCell(ci);
                            String cellValue = getCellValueAsString(cell);

                            Cell pdfCell = new Cell().add(new Paragraph(cellValue).setFontSize(9));
                            if (firstRow) {
                                pdfCell.setBackgroundColor(new com.itextpdf.kernel.colors.DeviceRgb(70, 130, 180))
                                       .setFontColor(ColorConstants.WHITE)
                                       .setBold();
                            } else if (rowCount % 2 == 0) {
                                pdfCell.setBackgroundColor(new com.itextpdf.kernel.colors.DeviceRgb(240, 248, 255));
                            }
                            table.addCell(pdfCell);
                        }
                        firstRow = false;
                        rowCount++;
                    }
                    pdfDocument.add(table);
                }
            }
            workbook.close();
        }
    }

    private static String getCellValueAsString(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    return new SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()).format(date);
                }
                double val = cell.getNumericCellValue();
                if (val == Math.floor(val) && !Double.isInfinite(val)) {
                    return String.valueOf((long) val);
                }
                return String.valueOf(val);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (Exception e) {
                    return cell.getCellFormula();
                }
            case BLANK:
            default:
                return "";
        }
    }

    // ─── POWERPOINT PPTX → PDF ─────────────────────────────────────────────

    private static void convertPptxToPdf(Context context, Uri uri, File output) throws Exception {
        try (InputStream is = context.getContentResolver().openInputStream(uri);
             XMLSlideShow ppt = new XMLSlideShow(is);
             PdfWriter writer = new PdfWriter(output);
             PdfDocument pdfDoc = new PdfDocument(writer);
             Document pdfDocument = new Document(pdfDoc)) {

            List<XSLFSlide> slides = ppt.getSlides();
            for (int i = 0; i < slides.size(); i++) {
                XSLFSlide slide = slides.get(i);

                // Encabezado de diapositiva
                pdfDocument.add(new Paragraph("Diapositiva " + (i + 1))
                    .setBold().setFontSize(13).setFontColor(ColorConstants.DARK_GRAY)
                    .setBackgroundColor(new com.itextpdf.kernel.colors.DeviceRgb(230, 230, 250))
                    .setMarginBottom(6).setMarginTop(i > 0 ? 15 : 0));

                // Extraer texto de cada forma
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape textShape = (XSLFTextShape) shape;
                        String text = textShape.getText();
                        if (text != null && !text.trim().isEmpty()) {
                            // El primer texto suele ser el título
                            boolean isTitle = slide.getShapes().indexOf(shape) == 0;
                            Paragraph p = new Paragraph(text)
                                .setFontSize(isTitle ? 14 : 11)
                                .setMarginBottom(4);
                            if (isTitle) p.setBold();
                            pdfDocument.add(p);
                        }
                    }
                }
            }
        }
    }

    private static void convertPptToPdf(Context context, Uri uri, File output) throws Exception {
        // Para .ppt antiguo
        try (InputStream is = context.getContentResolver().openInputStream(uri)) {
            org.apache.poi.hslf.usermodel.HSLFSlideShow ppt =
                new org.apache.poi.hslf.usermodel.HSLFSlideShow(is);

            try (PdfWriter writer = new PdfWriter(output);
                 PdfDocument pdfDoc = new PdfDocument(writer);
                 Document pdfDocument = new Document(pdfDoc)) {

                List<org.apache.poi.hslf.usermodel.HSLFSlide> slides = ppt.getSlides();
                for (int i = 0; i < slides.size(); i++) {
                    org.apache.poi.hslf.usermodel.HSLFSlide slide = slides.get(i);
                    pdfDocument.add(new Paragraph("Diapositiva " + (i + 1))
                        .setBold().setFontSize(13).setMarginBottom(6).setMarginTop(i > 0 ? 15 : 0));

                    for (org.apache.poi.hslf.usermodel.HSLFShape shape : slide.getShapes()) {
                        if (shape instanceof org.apache.poi.hslf.usermodel.HSLFTextShape) {
                            String text = ((org.apache.poi.hslf.usermodel.HSLFTextShape) shape).getText();
                            if (text != null && !text.trim().isEmpty()) {
                                pdfDocument.add(new Paragraph(text).setFontSize(11).setMarginBottom(3));
                            }
                        }
                    }
                }
            }
            ppt.close();
        }
    }
}
