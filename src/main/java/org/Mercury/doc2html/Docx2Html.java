package org.Mercury.doc2html;

import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

public class Docx2Html {
    public static void convert(String sourcePath, String destPath, final String imagePath) throws IOException {
        File targetFile = new File(sourcePath);
        OutputStreamWriter ows = null;

        try{
            XWPFDocument xwpfDocument = new XWPFDocument(new FileInputStream(sourcePath));
            XHTMLOptions options = XHTMLOptions.create();
            //存放图片的文件夹
            options.setExtractor(new FileImageExtractor(new File(imagePath)));
            //html中图片的路径
            options.URIResolver(new BasicURIResolver(imagePath));
            ows = new OutputStreamWriter(new FileOutputStream(destPath), "utf-8");
            XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
            xhtmlConverter.convert(xwpfDocument, ows, options);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if (ows != null) {
                ows.close();
            }
        }
    }
}
