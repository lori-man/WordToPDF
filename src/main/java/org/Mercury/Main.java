package org.Mercury;

import com.sun.xml.internal.messaging.saaj.util.ByteOutputStream;
import org.Mercury.doc2pdf.FileConverter;
import org.Mercury.docx2pdf.DocxToPDFConverter;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.ByteOrder;

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
    public void testdoc2PDF4() throws Exception {
        String sourcePath = "D:\\test\\20200624103913.doc";
        String destPath = "D:\\test\\test.html";
        String destPath2 = "D:\\test\\test.xhtml";
        String imagePath = "D:\\test";
        String pdfPath = "D:\\test\\test.pdf";
        FileInputStream fileInputStream = new FileInputStream(sourcePath);
        FileOutputStream fileOutputStream = new FileOutputStream(destPath);

        byte[] bytes = new byte[2048];
        int ch;
        while ((ch=fileInputStream.read(bytes)) != -1) {
            fileOutputStream.write(bytes, 0, ch);
        }
        FileConverter fileConverter = new FileConverter();
        fileConverter.parseToXhtml(destPath, destPath2);
        FileConverter.standardHTMLByTemplate(destPath2);
        fileConverter.convertHtmlToPdf(destPath2, pdfPath);

    }

    @Test
    public void testdoc2PDF5() throws Exception {
        String destPath = "D:\\test\\hello23456.html";
        String destPath2 = "D:\\test\\hello.xhtml";
        String imagePath = "D:\\test";
        String pdfPath = "D:\\test\\hello.pdf";
        FileConverter fileConverter = new FileConverter();
        fileConverter.parseToXhtml(destPath, destPath2);
        FileConverter.standardHTMLByTemplate(destPath2);
        fileConverter.convertHtmlToPdf(destPath2, pdfPath);

    }
    @Test
    public void testdocx2PDF3() throws Exception {
        String sourcePath = "D:\\test\\小程序接口.docx";
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

    @Test
    public void testdoc2PDF2() throws Exception {
        String sourcePath = "D:\\test\\hello1.doc";
        String destPath = "D:\\test\\hello.html";
        String imagePath = "D:\\test";
        String pdfPath = "D:\\test\\hello.pdf";
        String fontPath = "D:\\test\\simsun.ttf";
        FileInputStream fileInputStream = new FileInputStream(sourcePath);
        word2PDF.docToPDF(fileInputStream, destPath, pdfPath,fontPath);
    }

    @Test
    public void getUrl() throws IOException {
        BufferedInputStream bis = null;
        HttpURLConnection urlconnection = null;
        URL url = null;
        url = new URL("http://ckmro.oss-cn-shanghai.aliyuncs.com/ckmro/pdf/20200624/1592967000477.pdf");

        urlconnection = (HttpURLConnection) url.openConnection();
        urlconnection.connect();
        bis = new BufferedInputStream(urlconnection.getInputStream());
        byte[] bytes = new byte[2048];
        int ch;
        ByteOutputStream bos = new ByteOutputStream();
        while ((ch=bis.read(bytes)) != -1) {
            bos.write(bytes, 0, ch);
        }
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\test.pdf");
        fileOutputStream.write(bos.getBytes());
        System.out.println("file type:"+HttpURLConnection.guessContentTypeFromStream(bis));

    }
}


