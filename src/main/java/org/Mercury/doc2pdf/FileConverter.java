package org.Mercury.doc2pdf;

import com.lowagie.text.pdf.BaseFont;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.springframework.util.FileCopyUtils;
import org.w3c.dom.Document;
import org.w3c.tidy.Tidy;
import org.xhtmlrenderer.pdf.ITextFontResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class FileConverter {
    /**
     * word文件转成html文件
     *
     * @param sourceFilePath     src word文件路径
     * @param targetFilePosition 转化后生成的html文件路径
     * @param imagePath          图片路径
     * @throws Exception
     */
    public void wordToHtml(String sourceFilePath, String targetFilePosition, String imagePath) throws Exception {
        String suffix = sourceFilePath.substring(sourceFilePath.lastIndexOf(".", sourceFilePath.length()));
        switch (suffix) {
            case ".docx":
                docxToHtml(sourceFilePath, targetFilePosition, imagePath);
                break;
            case ".doc":
                docToHtml(sourceFilePath, targetFilePosition, imagePath);
                break;
            default:
                throw new RuntimeException("文件格式不正确");
        }
    }

    /**
     * doc转换为html
     */
    private void docToHtml(String sourceFilePath, String targetFilePosition, String imagePath) throws Exception {
        FileInputStream fileInputStream = new FileInputStream(sourceFilePath);
        HWPFDocument wordDocument = new HWPFDocument(fileInputStream);
        Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document);
        // 保存图片，并返回图片的相对路径
        wordToHtmlConverter.setPicturesManager(new MyPicturesManager(imagePath));
        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(new File(targetFilePosition));
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
    }



    /**
     * docx转换为html
     */
    private void docxToHtml(String sourceFilePath, String targetFileName,String imagePath) throws Exception {
        try (XWPFDocument document = new XWPFDocument(new FileInputStream(sourceFilePath));
             OutputStreamWriter outputStreamWriter = new OutputStreamWriter(new FileOutputStream(targetFileName), "UTF-8")) {
            XHTMLOptions options = XHTMLOptions.create();
            // 存放图片的文件夹
            options.setExtractor(new FileImageExtractor(new File(imagePath)));
            // html中图片的路径
            options.URIResolver(new BasicURIResolver(imagePath));
            XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
            xhtmlConverter.convert(document, outputStreamWriter, options);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 移动图片到指定路径
     */
    public  void changeImageUrl(String sourceFilePath,String targetFilePosition) throws IOException {
        try (FileInputStream fis = new FileInputStream(sourceFilePath);
             BufferedInputStream bufis = new BufferedInputStream(fis);

             FileOutputStream fos = new FileOutputStream(targetFilePosition);
             BufferedOutputStream bufos = new BufferedOutputStream(fos);) {

            int len = 0;
            while ((len = bufis.read()) != -1) {
                bufos.write(len);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * html文件解析成xhtml，变成标准的html文件
     * @param f_in 源html文件路径
     * @param outfile 输出后xhtml的文件路径
     */
    public boolean parseToXhtml(String f_in, String outfile) {
        boolean bo = false;
        String basil = null;
        try (FileInputStream fis = new FileInputStream(f_in);
             ByteArrayOutputStream bos = new ByteArrayOutputStream();) {
            int ch;
            while ((ch = fis.read()) != -1) {
                bos.write(ch);
            }
            byte[] bs = bos.toByteArray();
            bos.close();
            String hope_utf_8 = new String(bs, "UTF-8");// 默认是utf-8
            byte[] hope_b = hope_utf_8.getBytes();
            basil = new String(hope_b, "UTF-8");// 在此处可转换成其他字符类型
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 将生成的xhtml写入
        try (ByteArrayInputStream stream = new ByteArrayInputStream(basil.getBytes());
             ByteArrayOutputStream tidyOutStream = new ByteArrayOutputStream();
             DataOutputStream to = new DataOutputStream(new FileOutputStream(outfile));) {
            Tidy tidy = new Tidy();
            tidy.setInputEncoding("UTF-8");
            tidy.setQuiet(true);
            tidy.setOutputEncoding("UTF-8");
            tidy.setShowWarnings(true); // 不显示警告信息
            tidy.setIndentContent(true);//
            tidy.setSmartIndent(true);
            tidy.setIndentAttributes(false);
            tidy.setWraplen(1024); // 多长换行
            // 输出为xhtml
            tidy.setXHTML(true);
            tidy.setErrout(new PrintWriter(System.out));
            tidy.parse(stream, tidyOutStream);
            tidyOutStream.writeTo(to);
            bo = true;
        }catch (IOException e) {
            e.printStackTrace();
        }
        return bo;
    }

    /**
     * xhtml文件转pdf文件
     */
    public boolean convertHtmlToPdf(String inputFile, String outputFile) {
        try (OutputStream os = new FileOutputStream(outputFile);
             InputStream resourceAsStream = this.getClass().getClassLoader().getResourceAsStream("simsun.ttf");) {
            ITextRenderer renderer = new ITextRenderer();
            String url = new File(inputFile).toURI().toURL().toString();

            // 解决中文支持问题
            ITextFontResolver fontResolver = renderer.getFontResolver();
            String fontSimSun = "http://static.test.ckmro.com:9035/static/simsun.ttf";

            File file = new File("D:\\workspace\\simsun.ttf");
            if (!file.exists()) {
                file.createNewFile();
            }
            FileCopyUtils.copy(resourceAsStream, new FileOutputStream(file));
            fontSimSun = "D:\\workspace\\simsun.ttf";

            fontResolver.addFont(fontSimSun, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            // 解决图片的相对路径问题
//            renderer.getSharedContext().setBaseURL("D:\\test\\imageTmp");
            renderer.setDocument(url);
            renderer.layout();
            renderer.createPDF(os);
            os.flush();
            return true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;
    }

    /*
     * xhtml转成标准html文件
     * targetHtml:要处理的html文件路径
     */
    public static void standardHTML(String targetHtml) throws IOException {
        File f = new File(targetHtml);
        org.jsoup.nodes.Document doc = Jsoup.parse(f, "UTF-8");

        doc.select("meta").removeAttr("name");
        doc.select("meta").attr("content", "text/html; charset=UTF-8");
        doc.select("meta").attr("http-equiv", "Content-Type");
        doc.select("meta").html("&nbsp");
        doc.select("img").html("&nbsp");
        doc.select("style").attr("mce_bogus", "1");
        doc.select("body").attr("style", "font-family: SimSun");
        doc.select("html").before("<?xml version='1.0' encoding='UTF-8'>");
        Elements style = doc.getElementsByTag("style");
        style.get(0).append("body {\n" +
                "            font-family: SimSun;}\n" +
                "@page { size: A4 }");

        setTagFont(doc);

        /*
         * Jsoup只是解析，不能保存修改，所以要在这里保存修改。
         */
        FileOutputStream fos = new FileOutputStream(f, false);
        OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
        osw.write(doc.html());
        osw.close();
    }

    /*
     * xhtml转成标准html文件
     * targetHtml:要处理的html文件路径
     */
    public static void standardHTMLByTemplate(String targetHtml) throws IOException {
        File f = new File(targetHtml);
        org.jsoup.nodes.Document doc = Jsoup.parse(f, "UTF-8");

        doc.select("meta").removeAttr("name");
        doc.select("meta").attr("content", "text/html; charset=UTF-8");
        doc.select("meta").attr("http-equiv", "Content-Type");
        doc.select("meta").html("&nbsp");
        doc.select("img").html("&nbsp");
        doc.select("style").attr("mce_bogus", "1");
        doc.select("style").html("");
        doc.select("body").attr("style", "font-family: SimSun");
        doc.select("html").before("<?xml version='1.0' encoding='UTF-8'>");
        Elements style = doc.getElementsByTag("style");
        style.get(0).append("span{" +
                "            color:black}" +
                "body {\n" +
                "            font-family: SimSun;}\n" +
                "@page { size: A4;padding-top: 60pt }" +
                "div.header {\n" +
                "    display: block; text-align: center; \n" +
                "    position: running(header);\n" +
                "}\n" +
                "div.footer {\n" +
                "    display: block; text-align: right;\n" +
                "    position: running(footer);\n" +
                "}\n" +
                "#pagenumber:before {\n" +
                "    content: counter(page);\n" +
                "}\n" +
                "#pagecount:before {\n" +
                "    content: counter(pages);\n" +
                "}" +
                "div.content {page-break-after: always;}\n" +
                "@page {\n" +
                "     @top-center { content: element(header) }\n" +
                "}\n" +
                "@page { \n" +
                "    @bottom-center { content: element(footer) }\n" +
                "}");

        setTagFont(doc);

        Elements divs = doc.getElementsByTag("div");
        Element element = divs.get(1);
        element.attr("align", "left");
        Elements tables = doc.select("table");
        Element table1 = tables.get(1);
        table1.attr("border", "1px");
        Element element3 = divs.get(0);
        Elements element3ElementsByDiv = element3.getElementsByTag("div");
        Element element4 = element3ElementsByDiv.get(0);
        element4.attr("align", "left");

        divs.get(0).before("    <div class='header'><img style=\"height:80pt;margin-top: 20pt\" src='http://static.ckmro.com:8082/static/contract.files/image001.png'/></div>\n" +
                "    <div class='footer'><span id=\"pagenumber\"></span></div>");

        //处理legal表格
        Element legalTable = tables.get(2);
        Elements rows = legalTable.getElementsByTag("tr");
        legalTable.attr("width", "100%");
        for (Element row : rows) {
//            row.attr("text-align", "left");
            Elements allElements = row.getAllElements();
            allElements.removeAttr("width");
            Elements tds = row.getElementsByTag("td");
            tds.attr("width", "50%");
            boolean b = allElements.stream().anyMatch(smallEle -> {
                return "&nbsp;".equals(smallEle.html());
            });
            if (b) {
                row.remove();
            }
        }
        Element lastDiv = divs.get(divs.size() - 1);
//        lastDiv.attr("style", "font-family: SimSun;padding-left: 45pt");
        lastDiv.attr("align", "left");
        /*
         * Jsoup只是解析，不能保存修改，所以要在这里保存修改。
         */
        FileOutputStream fos = new FileOutputStream(f, false);
        OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
        osw.write(doc.html());
        osw.close();
    }

    private static void setTagFont(org.jsoup.nodes.Document doc) {
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
    }

    private class MyPicturesManager implements PicturesManager {
        private final String imagePath;

        private MyPicturesManager(String imagePath) {
            this.imagePath = imagePath;
        }
        @Override
        public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches, float heightInches) {
            try (FileOutputStream out = new FileOutputStream(imagePath + "\\" + suggestedName)) {
                out.write(content);
            } catch (IOException e) {
                e.printStackTrace();
            }
            return imagePath + "\\" + suggestedName;
        }
    }

}
//https://cloud.tencent.com/developer/ask/44759