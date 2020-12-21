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
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class FileConverter {
    /*
     * word文件转成html文件
     * sourceFilePath:源word文件路径
     * targetFilePosition:转化后生成的html文件路径
     */
    public void wordToHtml(String sourceFilePath, String targetFilePosition,String imagePath) throws Exception {
        if (".docx".equals(sourceFilePath.substring(sourceFilePath.lastIndexOf(".", sourceFilePath.length())))) {
            docxToHtml(sourceFilePath, targetFilePosition,imagePath);
        } else if (".doc".equals(sourceFilePath.substring(sourceFilePath.lastIndexOf(".", sourceFilePath.length())))) {
            docToHtml(sourceFilePath, targetFilePosition,imagePath);
        } else {
            throw new RuntimeException("文件格式不正确");
        }

    }

    /*
     * doc转换为html
     * sourceFilePath:
     * targetFilePosition:生成的html文件路源word文件路径径
     */

    private void docToHtml(String sourceFilePath, String targetFilePosition, String imagePath) throws Exception {
        FileInputStream fileInputStream = new FileInputStream(sourceFilePath);
        HWPFDocument wordDocument = new HWPFDocument(fileInputStream);
        Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document);
        // 保存图片，并返回图片的相对路径
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            @Override
            public String savePicture(byte[] content, PictureType pictureType, String name, float width, float height) {
                try {
                    FileOutputStream out = new FileOutputStream(imagePath + "\\" + name);
                    out.write(content);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                return imagePath + "\\" + name;
            }
        });
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


    /*
     * docx转换为html
     * sourceFilePath:源word文件路径
     * targetFileName:生成的html文件路径
     */

    private void docxToHtml(String sourceFilePath, String targetFileName,String imagePath) throws Exception {
        OutputStreamWriter outputStreamWriter = null;
        try {
            XWPFDocument document = new XWPFDocument(new FileInputStream(sourceFilePath));
            XHTMLOptions options = XHTMLOptions.create();
            // 存放图片的文件夹
            options.setExtractor(new FileImageExtractor(new File(imagePath)));
            // html中图片的路径
            options.URIResolver(new BasicURIResolver(imagePath));
            outputStreamWriter = new OutputStreamWriter(new FileOutputStream(targetFileName), "UTF-8");
            XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
            xhtmlConverter.convert(document, outputStreamWriter, options);
        } finally {
            if (outputStreamWriter != null) {
                outputStreamWriter.close();
            }
        }
    }


    /*
    移动图片到指定路径
    sourceFilePath:原始路径
    targetFilePosition:移动后存放的路径
    */

    public  void changeImageUrl(String sourceFilePath,String targetFilePosition) throws IOException {
        FileInputStream fis = new FileInputStream(sourceFilePath);
        BufferedInputStream bufis = new BufferedInputStream(fis);

        FileOutputStream fos = new FileOutputStream(targetFilePosition);
        BufferedOutputStream bufos = new BufferedOutputStream(fos);
        int len = 0;
        while ((len = bufis.read()) != -1) {
            bufos.write(len);
        }
        bufis.close();
        bufos.close();
    }



    /*
     * html文件解析成xhtml，变成标准的html文件
     * f_in:源html文件路径
     * outfile: 输出后xhtml的文件路径
     */
    public boolean parseToXhtml(String f_in, String outfile) {
        boolean bo = false;
        ByteArrayOutputStream tidyOutStream = null; // 输出流
        FileInputStream fis = null;
        ByteArrayOutputStream bos = null;
        ByteArrayInputStream stream = null;
        DataOutputStream to = null;
        try {
            // Reader reader;
            fis = new FileInputStream(f_in);
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
            tidy.setShowWarnings(true); // 不显示警告信息
            tidy.setIndentContent(true);//
            tidy.setSmartIndent(true);
            tidy.setIndentAttributes(false);
            tidy.setWraplen(1024); // 多长换行
            // 输出为xhtml
            tidy.setXHTML(true);
            tidy.setErrout(new PrintWriter(System.out));
            tidy.parse(stream, tidyOutStream);
            to = new DataOutputStream(new FileOutputStream(outfile));// 将生成的xhtml写入
            tidyOutStream.writeTo(to);
            bo = true;
        } catch (Exception ex) {
            System.out.println(ex.toString());
            ex.printStackTrace();
            return bo;
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
            System.gc();
        }
        return bo;
    }

    /*
     * xhtml文件转pdf文件
     * inputFile:xhtml源文件路径
     * outputFile:输出的pdf文件路径
     * imagePath:图片的存放路径   例如(file:/D:/test)
     */
    public boolean convertHtmlToPdf(String inputFile, String outputFile) throws Exception {
        OutputStream os = new FileOutputStream(outputFile);
        ITextRenderer renderer = new ITextRenderer();
        String url = new File(inputFile).toURI().toURL().toString();

        // 解决中文支持问题
        ITextFontResolver fontResolver = renderer.getFontResolver();
        //D:\test\simsun.ttf
/*        String resource = getClass().getClassLoader().getResource("simsun.ttf").getPath();
        fontResolver.addFont(resource, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);*/
//        fontResolver.addFont("http://static.test.ckmro.com:9035/static/simsun.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
/*        System.out.println("字体处理开始:"+new Date());
        fontResolver.addFont("http://static.ckmro.com:8082/static/simsun.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        System.out.println("字体处理结束:"+new Date());*/
        String fontSimSun = "http://static.test.ckmro.com:9035/static/simsun.ttf";
        try{
            InputStream resourceAsStream = this.getClass().getClassLoader().getResourceAsStream("simsun.ttf");
            File file = new File("D:\\workspace\\simsun.ttf");
            if (!file.exists()) {
                file.createNewFile();
            }
            FileCopyUtils.copy(resourceAsStream, new FileOutputStream(file));
            fontSimSun = "D:\\workspace\\simsun.ttf";
        }catch (FileNotFoundException e) {
            System.out.println("获取File失败");
        } catch (IOException e) {
            System.out.println("文件拷贝失败");
        }
        fontResolver.addFont(fontSimSun, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        // 解决图片的相对路径问题
        renderer.getSharedContext().setBaseURL("imagePath");
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
        Elements select = doc.select("body > *");
        List<Element> list = new ArrayList<>(select);
        while (list.size() > 0) {
            Element element = list.get(0);
            element.attr("style", "font-family: SimSun;");//;
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
                "@page { size: A4 }");
        Elements select = doc.select("body > *");
        List<Element> list = new ArrayList<>(select);
        while (list.size() > 0) {
            Element element = list.get(0);
            element.attr("style", "font-family: SimSun;");//;
            Elements select1 = element.getAllElements();
            select1.remove(element);
            list.addAll(select1);
            list.remove(element);
        }
        Element element = doc.select("div").get(1);
        element.attr("align", "left");
        Elements tables = doc.select("table");
        Element table1 = tables.get(1);
        table1.attr("border", "1px");
        /*
         * Jsoup只是解析，不能保存修改，所以要在这里保存修改。
         */
        FileOutputStream fos = new FileOutputStream(f, false);
        OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
        osw.write(doc.html());
        osw.close();
    }

}
