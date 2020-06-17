package org.Mercury;

import org.Mercury.docx4j.Docx2PDF;
import org.Mercury.docxl2pdf.FileConverter;
import org.junit.jupiter.api.Test;

public class Main {

    @Test
    public void testdoc2PDF() throws Exception {
        String sourcePath = "D:\\test\\hello.doc";
        String destPath = "D:\\test\\hello.html";
        String destPath2 = "D:\\test\\hello.xhtml";
        String imagePath = "D:\\test";
        String pdfPath = "D:\\test\\hello.pdf";
        FileConverter fileConverter = new FileConverter();
        fileConverter.wordToHtml(sourcePath, destPath,imagePath);
        fileConverter.parseToXhtml(destPath, destPath2);
        FileConverter.standardHTML(destPath2);
        fileConverter.convertHtmlToPdf(destPath2, pdfPath);

    }
    @Test
    public void testdocx2PDF() throws Exception {
        String sourcePath = "D:\\test\\hello.docx";
        String destPath = "D:\\test\\hello.xhtml";
        String imagePath = "D:\\test";
        String pdfPath = "D:\\test\\hello.pdf";

        FileConverter fileConverter = new FileConverter();
        fileConverter.wordToHtml(sourcePath, destPath,imagePath);
        fileConverter.parseToXhtml(destPath, destPath);
        FileConverter.standardHTML(destPath);
        fileConverter.convertHtmlToPdf(destPath, pdfPath);

    }

    @Test
    public void testHtml2PDF2() throws Exception {
        String pdfPath = "D:\\test\\hello.pdf";
        String sourcePath = "D:\\test\\hello.docx";
        Docx2PDF.convertDocxToPDF(sourcePath, pdfPath);
    }


    @Test
    public void testdocx2PDF2() throws Exception {
        String sourcePath = "D:\\test\\example.html";
        String destPath = "D:\\test\\hello.xhtml";
        String imagePath = "D:\\test";
        String pdfPath = "D:\\test\\hello.pdf";

        FileConverter fileConverter = new FileConverter();
        fileConverter.parseToXhtml(sourcePath, destPath);
        FileConverter.standardHTML(destPath);
        fileConverter.convertHtmlToPdf(destPath, pdfPath);

    }
}


