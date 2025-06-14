package com.pdf.excel;

import java.awt.Rectangle; // Import Rectangle for bounding box
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.Word;

@SpringBootApplication
public class PdfToExcelConverterApplication {

	/*
	 * public static void main(String[] args) {
	 * SpringApplication.run(PdfToExcelConverterApplication.class, args); }
	 */
	
	// IMPORTANT: Set this to the path where your Tesseract-OCR installation is located.
    // For example: "C:/Program Files/Tesseract-OCR" on Windows, or "/usr/local/share/tessdata"
    // if you installed Tesseract via Homebrew on macOS (and set TESSDATA_PREFIX env var).
    // The 'tessdata' folder with language files (e.g., eng.traineddata) must be accessible.
    private static final String TESSDATA_PATH = "C:/Program Files/Tesseract-OCR/tessdata"; // <--- CHANGE THIS PATH

    public static void main(String[] args) {
    	
    	//SpringApplication.run(PdfToExcelConverterApplication.class, args);
    	
        if (args.length < 2) {
            System.out.println("Usage: java -jar pdf-to-excel.jar <inputPdfPath> <outputExcelPath>");
            System.out.println("Example: java -jar pdf-to-excel.jar input.pdf output.xlsx");
            return;
        }

        String inputPdfPath = args[0];
        String outputExcelPath = args[1];

        // Ensure Tesseract data path is accessible for Tess4J
        File tessDataDir = new File(TESSDATA_PATH);
        if (!tessDataDir.exists() || !tessDataDir.isDirectory()) {
            System.err.println("Error: TESSDATA_PATH not found or is not a directory: " + TESSDATA_PATH);
            System.err.println("Please install Tesseract OCR and set TESSDATA_PATH correctly.");
            System.err.println("Refer to the instructions in the pom.xml and this file.");
            return;
        }

        try {
            convertPdfImageToExcel(inputPdfPath, outputExcelPath);
            System.out.println("PDF conversion to Excel completed successfully!");
            System.out.println("Output saved to: " + outputExcelPath);
        } catch (Exception e) {
            System.err.println("An error occurred during conversion: " + e.getMessage());
            e.printStackTrace();
        }
    }

    public static void convertPdfImageToExcel(String pdfPath, String excelPath) throws IOException {
        PDDocument document = null;
        Workbook workbook = new XSSFWorkbook();
        Path tempDir = null;

        try {
            document = PDDocument.load(new File(pdfPath));
            PDFRenderer pdfRenderer = new PDFRenderer(document);

            // Create a temporary directory for image files
            tempDir = Files.createTempDirectory("pdf_images_");
            System.out.println("Temporary image directory created at: " + tempDir.toAbsolutePath());

            ITesseract tesseract = new Tesseract();
            // It's crucial to set the datapath correctly for Tesseract to find language files.
            tesseract.setDatapath(TESSDATA_PATH);
            tesseract.setLanguage("eng"); // Set the language, e.g., "eng" for English

            System.out.println("Starting PDF page rendering and OCR...");

            // Process each page
            for (int i = 0; i < document.getNumberOfPages(); i++) {
                System.out.println("Processing Page " + (i + 1) + " of " + document.getNumberOfPages());

                // Render PDF page to a high-resolution image (300 DPI is good for OCR)
                BufferedImage bim = pdfRenderer.renderImageWithDPI(i, 300);
                // Saving the image to a temporary file is not strictly necessary for getWords(BufferedImage, int)
                // but kept here for debugging/inspection purposes.
                Path tempImagePath = tempDir.resolve("page_" + (i + 1) + ".png");
                ImageIO.write(bim, "png", tempImagePath.toFile());
                System.out.println(" - Rendered page " + (i + 1) + " to " + tempImagePath.toAbsolutePath());

                // Perform OCR on the rendered image
                List<Word> words = null;
                try {
                    // getWords method provides bounding box for each word
                    // Using ITesseract.RIL_WORD to get words at the word level
                    words = tesseract.getWords(bim, 3);
                    System.out.println(" - OCR completed for page " + (i + 1) + ". Found " + words.size() + " words.");
                } catch (Exception e) {
                    System.err.println(" - Error during OCR for page " + (i + 1) + ": " + e.getMessage());
                    // Continue to next page even if OCR fails for one
                    continue;
                }

                // Create a new sheet for each PDF page
                Sheet sheet = workbook.createSheet("Page " + (i + 1));

                // Heuristic: Estimate pixel to Excel cell conversion ratio.
                // These values are highly approximate and might need tuning based on DPI and desired Excel look.
                // A typical Excel default row height is about 15 points, and column width is 8.43 chars.
                // If 300 DPI image, 1 inch = 300 pixels.
                // Let's assume 1 Excel row height ~ 20 pixels and 1 Excel column width ~ 50 pixels at 300 DPI.
                final double PIXEL_TO_ROW_RATIO = 20.0; // Pixels per "virtual" Excel row unit
                final double PIXEL_TO_COL_RATIO = 50.0; // Pixels per "virtual" Excel column unit

                // Sort words by Y-coordinate then X-coordinate for a more natural reading order
                words.sort(Comparator.<Word>comparingDouble(w -> w.getBoundingBox().getY()) // Correctly accessing Y from BoundingBox
                        .thenComparingDouble(w -> w.getBoundingBox().getX())); // Correctly accessing X from BoundingBox

                // Keep track of max row and column for setting dimensions
                int maxRow = 0;
                int maxCol = 0;

                // Place OCR'd text into Excel cells
                for (Word word : words) {
                    Rectangle boundingBox = word.getBoundingBox(); // Get the bounding box
                    int rowNum = (int) (boundingBox.getY() / PIXEL_TO_ROW_RATIO); // Use Y from boundingBox
                    int colNum = (int) (boundingBox.getX() / PIXEL_TO_COL_RATIO); // Use X from boundingBox

                    Row row = sheet.getRow(rowNum);
                    if (row == null) {
                        row = sheet.createRow(rowNum);
                    }

                    Cell cell = row.getCell(colNum);
                    if (cell == null) {
                        cell = row.createCell(colNum);
                    }

                    // Append text if cell already contains something (simple merging attempt)
                    String currentCellValue = "";
                    try {
                        currentCellValue = cell.getStringCellValue();
                    } catch (IllegalStateException e) {
                        // Cell might contain a numeric value or be empty
                        // For simplicity, if it's not a string, treat as empty for appending
                        currentCellValue = "";
                    }

                    if (!currentCellValue.isEmpty()) {
                        cell.setCellValue(currentCellValue + " " + word.getText());
                    } else {
                        cell.setCellValue(word.getText());
                    }

                    // Update max row and col
                    if (rowNum > maxRow) maxRow = rowNum;
                    if (colNum > maxCol) maxCol = colNum;

                    // Attempt to adjust column width based on content.
                    // This is a basic approach; `autoSizeColumn` might be better after all data is in.
                    // Or set a fixed width based on a heuristic.
                    // sheet.setColumnWidth(colNum, sheet.getColumnWidth(colNum) + (word.getText().length() * 256)); // POI units
                }

                // Auto-size columns for better readability (can be slow for many columns)
                // This should ideally be done AFTER all content is placed in the sheet.
                for (int col = 0; col <= maxCol; col++) {
                    sheet.autoSizeColumn(col);
                }

                // Adjust row height based on content or a fixed value for consistency
                // For simplicity, we'll let Excel handle row heights based on auto-sizing columns
                // or just keep default. Setting row height explicitly requires calculating
                // font metrics, which is complex.
            }

            // Write the workbook to the output Excel file
            try (FileOutputStream fileOut = new FileOutputStream(excelPath)) {
                workbook.write(fileOut);
            }
        } catch (Exception e) {
            System.err.println("Error during PDF to Excel conversion: " + e.getMessage());
            throw new IOException("Conversion failed", e);
        } finally {
            if (document != null) {
                try {
                    document.close();
                } catch (IOException e) {
                    System.err.println("Error closing PDF document: " + e.getMessage());
                }
            }
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    System.err.println("Error closing Excel workbook: " + e.getMessage());
                }
            }
            // Clean up temporary image files
            if (tempDir != null && Files.exists(tempDir)) {
                System.out.println("Cleaning up temporary image directory: " + tempDir.toAbsolutePath());
                try {
                    Files.walk(tempDir)
                            .sorted(Comparator.reverseOrder())
                            .map(Path::toFile)
                            .forEach(File::delete);
                } catch (IOException e) {
                    System.err.println("Error deleting temporary files: " + e.getMessage());
                }
            }
        }
    }

}
