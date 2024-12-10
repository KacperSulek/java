package org.example;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("PDF Files", "pdf");

        chooser.setFileFilter(filter);

        int returnVal = chooser.showOpenDialog(null);

        if (returnVal == JFileChooser.APPROVE_OPTION) {
            String pdfPath = chooser.getSelectedFile().getPath(); // Path to your PDF file

            filter = new FileNameExtensionFilter("Word", "doc");

            chooser.setFileFilter(filter);

            chooser.showSaveDialog(null);

            String docPath = chooser.getSelectedFile().getPath(); // Path to save the Word document

            try {
                convertPdfToDoc(pdfPath, docPath);
                System.out.println("Conversion completed. File saved at: " + docPath);
            } catch (IOException e) {
                System.err.println("An error occurred during conversion: " + e.getMessage());
            }
        }
    }

    public static void convertPdfToDoc(String pdfPath, String docPath) throws IOException {
        // Load the PDF document
        try (PDDocument pdfDocument = Loader.loadPDF(new File(pdfPath))) {
            // Extract text from the PDF
            PDFTextStripper pdfStripper = new PDFTextStripper();
            String text = pdfStripper.getText(pdfDocument);

            // Create a new Word document
            try (XWPFDocument wordDocument = new XWPFDocument();
                 FileOutputStream outputStream = new FileOutputStream(docPath)) {

                // Add the text to a Word paragraph
                XWPFParagraph paragraph = wordDocument.createParagraph();
                paragraph.createRun().setText(text);

                // Write the Word document to the file
                wordDocument.write(outputStream);
            }
        }
    }
}