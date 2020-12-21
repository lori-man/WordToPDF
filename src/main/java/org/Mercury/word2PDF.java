package org.Mercury;

import com.lowagie.text.pdf.BaseFont;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.dom.Document;
import org.w3c.tidy.Tidy;
import org.xhtmlrenderer.pdf.ITextFontResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class word2PDF {

    /*
     * doc转换为html
     * sourceFilePath:
     * targetFilePosition:生成的html文件路源word文件路径径
     */

    public static void docToPDF(InputStream sourceFile, String targetHtmlPath, String targetPDFPath, String fontPath) {
        FileInputStream fileInputStream = null;
        try {
            HWPFDocument wordDocument = new HWPFDocument(sourceFile);
            Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document);

            wordToHtmlConverter.processDocument(wordDocument);
            Document htmlDocument = wordToHtmlConverter.getDocument();
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(new File(targetHtmlPath));
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);

            byte[] bytes = parseToXhtml(targetHtmlPath);
            standardHTML(targetHtmlPath);
            convertHtmlToPdf(targetHtmlPath, targetPDFPath,fontPath);
        } catch (TransformerConfigurationException e) {
            e.printStackTrace();
        } catch (TransformerException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (fileInputStream != null) {
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /*
     * html文件解析成xhtml，变成标准的html文件
     */
    public static byte[] parseToXhtml(String targetPath) {
        ByteArrayOutputStream tidyOutStream = null; // 输出流
        ByteArrayOutputStream bos = null;
        ByteArrayInputStream stream = null;
        FileInputStream fis = null;
        DataOutputStream to = null;
        try {
            fis = new FileInputStream(targetPath);
            // Reader reader;
            bos = new ByteArrayOutputStream();
            int ch;
            while ((ch = fis.read()) != -1) {
                bos.write(ch);
            }
            byte[] bs = bos.toByteArray();
            bos.close();
            String hope_utf_8 = new String(bs, "UTF-8");// 默认是utf-8
            byte[] hope_b = hope_utf_8.getBytes();
            String basil = new String(hope_b, "UTF-8");// 在此处可转换成其他字符类型
            stream = new ByteArrayInputStream(basil.getBytes());
            tidyOutStream = new ByteArrayOutputStream();

            Tidy tidy = new Tidy();
            tidy.setInputEncoding("UTF-8");
            tidy.setQuiet(true);
            tidy.setOutputEncoding("UTF-8");
            tidy.setShowWarnings(true);
            tidy.setIndentContent(true);
            tidy.setSmartIndent(true);
            tidy.setIndentAttributes(false);
            tidy.setWraplen(1024);

            // 输出为xhtml
            tidy.setXHTML(true);
            tidy.setErrout(new PrintWriter(System.out));
            tidy.parse(stream, tidyOutStream);
            to = new DataOutputStream(new FileOutputStream(targetPath));
            tidyOutStream.writeTo(to);
            return tidyOutStream.toByteArray();
        } catch (Exception ex) {
            System.out.println(ex.toString());
            ex.printStackTrace();
            return null;
        } finally {
            try {
                if (to != null) {
                    to.close();
                }
                if (stream != null) {
                    stream.close();
                }
                if (fis != null) {
                    fis.close();
                }
                if (bos != null) {
                    bos.close();
                }
                if (tidyOutStream != null) {
                    tidyOutStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            System.gc();  //垃圾回收
        }
    }

    public static boolean convertHtmlToPdf(String inputFile, String outputFile,String fontPath) throws Exception {
        OutputStream os = new FileOutputStream(outputFile);
        ITextRenderer renderer = new ITextRenderer();
        String url = new File(inputFile).toURI().toURL().toString();


        // 解决中文支持问题
        ITextFontResolver fontResolver = renderer.getFontResolver();
        fontResolver.addFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

        renderer.setDocument(url);
        renderer.layout();
        renderer.createPDF(os);
        os.flush();
        os.close();
        return true;
    }

    /*
     * xhtml转成标准html文件
     * targetHtml:要处理的html文件路径
     */
    public static void standardHTML(String targetHtml){
        try{
            File f = new File(targetHtml);
            org.jsoup.nodes.Document doc = Jsoup.parse(f, "UTF-8");

            doc.select("meta").removeAttr("name");
            doc.select("meta").attr("content", "text/html; charset=UTF-8");
            doc.select("meta").attr("http-equiv", "Content-Type");
            doc.select("style").attr("mce_bogus", "1");
            doc.select("body").attr("style", "font-family: SimSun");
            doc.select("html").before("<?xml version='1.0' encoding='UTF-8'>");
            Elements style = doc.getElementsByTag("style");
            style.get(0).append("body {\n" +
                    "            font-family: SimSun;\n"  +
                    "@page{size: A4}");
            Elements select = doc.select("body > *");
            List<Element> list = new ArrayList<>(select);
            while (list.size() > 0) {
                Element element = list.get(0);
                element.attr("style", "font-family: SimSun;");
                Elements select1 = element.getAllElements();
                select1.remove(element);
                list.addAll(select1);
                list.remove(element);
            }
            /*
             * Jsoup只是解析，不能保存修改，所以要在这里保存修改。
             */
            FileOutputStream fos = new FileOutputStream(f, false);
            OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
            osw.write(doc.html());
            osw.close();
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
