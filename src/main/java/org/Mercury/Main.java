package org.Mercury;

import org.Mercury.doc2pdf.FileConverter;
import org.Mercury.docx2pdf.DocxToPDFConverter;
import org.Mercury.html2pdf.Html2PDF;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

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
    public void testdocx2PDF3() throws Exception {
        String sourcePath = "D:\\test\\hello.docx";
        String destPath = "D:\\test\\hello.xhtml";
        String imagePath = "D:\\test";
        String pdfPath = "D:\\test\\hello.pdf";

        File file = new File(sourcePath);
        FileInputStream fileInputStream = new FileInputStream(file);

        FileOutputStream fileOutputStream = new FileOutputStream(pdfPath);
        DocxToPDFConverter converter = new DocxToPDFConverter(fileInputStream, fileOutputStream, true, true);
        converter.convert();
    }

    @Test
    public void testdocx2PDF2() throws Exception {
        String sourcePath = "D:\\test\\hello.docx";
        String destPath = "D:\\test\\hello.xhtml";
        String imagePath = "D:\\test";
        String pdfPath = "D:\\test\\hello.pdf";

        File file = new File(sourcePath);
        FileInputStream fileInputStream = new FileInputStream(file);

        FileOutputStream fileOutputStream = new FileOutputStream(pdfPath);
        DocxToPDFConverter converter = new DocxToPDFConverter(fileInputStream, fileOutputStream, true, true);
        converter.convert();
    }

}


