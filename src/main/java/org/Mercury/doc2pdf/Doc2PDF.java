package com.mro.commons.utils.word2pdf;

import com.lowagie.text.pdf.BaseFont;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.springframework.util.Assert;
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
import java.util.Iterator;
import java.util.List;

public class Doc2PDF {

    /**
     * word转pdf
     *
     * @param in             需转换word的输入流
     * @param targetHtmlPath 生成html的位置
     * @param targetPDFPath  生成pdf的位置
     * @param fontPath       字体位置
     */
    public static void docToPDF(InputStream in, String targetHtmlPath, String targetXhtmlPath, String targetPDFPath, String fontPath) throws Exception {
        HWPFDocument wordDocument = new HWPFDocument(in);
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

        parseToXhtml(targetHtmlPath, targetXhtmlPath);
        standardHTML(targetXhtmlPath);
        convertHtmlToPdf(targetXhtmlPath, targetPDFPath, fontPath);
    }

    public static void docToPDFByTemplate(byte[] in, String targetHtmlPath, String targetXhtmlPath, String targetPDFPath, String fontPath) throws Exception {
        try(FileOutputStream fileOutputStream = new FileOutputStream(targetHtmlPath);) {
            fileOutputStream.write(in);
            parseToXhtml(targetHtmlPath,targetXhtmlPath);
            standardHTMLByTemplate(targetXhtmlPath);
            convertHtmlToPdf(targetXhtmlPath, targetPDFPath, fontPath);
        }
    }

    /**
     * html文件解析成xhtml，变成标准的html文件
     *
     * @param f_in    源html文件路径
     * @param outfile 输出后xhtml的文件路径
     */
    public static boolean parseToXhtml(String f_in, String outfile) throws Exception {
        boolean bo = false;
        String basil = null;
        try (FileInputStream fis = new FileInputStream(f_in);
             ByteArrayOutputStream bos = new ByteArrayOutputStream();) {
            int ch;
            while ((ch = fis.read()) != -1) {
                bos.write(ch);
            }
            byte[] bs = bos.toByteArray();
            String hope_utf_8 = new String(bs, "UTF-8");// 默认是utf-8
            byte[] hope_b = hope_utf_8.getBytes();
            basil = new String(hope_b, "UTF-8");// 在此处可转换成其他字符类型
        }

        Assert.notNull(basil, "parseToXhtml basil is null");

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
        }
        return bo;
    }

    /**
     * xhtml文件转pdf文件
     */
    public static boolean convertHtmlToPdf(String inputFile, String outputFile, String fontPath) throws Exception {
        try (OutputStream os = new FileOutputStream(outputFile);) {
            ITextRenderer renderer = new ITextRenderer();
            String url = new File(inputFile).toURI().toURL().toString();

            // 解决中文支持问题
            ITextFontResolver fontResolver = renderer.getFontResolver();
            fontResolver.addFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            renderer.setDocument(url);
            renderer.layout();
            renderer.createPDF(os);
            os.flush();
            return true;
        }
    }

    /**
     *  xhtml转成标准html文件
     * @param targetHtml 要处理的html文件路径
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

    /**
     * xhtml转成标准html文件
     * @param targetHtml 要处理的html文件路径
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
//                "@page { size: A4;padding-top: 60pt }" +
                "@page { size: A4;padding: 48pt;padding-top: 60pt }" +
                "div.header {\n" +
                "    padding-left:53pt;display: block; text-align: center; \n" +
                "    position: running(header);\n" +
                "}\n" +
                "p.MsoNormal, li.MsoNormal, div.MsoNormal {\n" +
                "    margin: 0cm;\n" +
                "    margin-bottom: .0001pt;\n" +
                "    line-height: 15.75pt;\n" +
                "    font-size: 10.5pt;\n" +
                "    font-family: 宋体;\n" +
                "}" +
                "div.footer {\n" +
                "    margin-bottom:10pt;display: block; text-align: right;\n" +
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

        divs.get(0).before("    <div class='header'><img style=\"height:65pt;margin-top: 20pt\" src='http://static.ckmro.com:8082/static/contract.files/image001.png'/></div>\n" +
                "    <div class='footer'><span id=\"pagenumber\"></span></div>");

        //处理字体大小
        Elements spans = doc.body().select("span");
        spans.attr("style", "font-family: SimSun;font-size:10.5pt");
        //采控网销售合同
        Iterator<Element> iterator = spans.iterator();
        Element first = iterator.next();
        Element second = iterator.next();
        second.attr("style", "font-family: SimSun;font-size:22.0pt");
//        //合同编号
        Element firstDiv = divs.get(0);
        Elements ps = firstDiv.select("p");
//        Element p1 = ps.get(1);
//        p1.attr("style","text-indent:240.0pt;line-height:4.0pt;text-autospace:\n" +
//                "none;vertical-align:bottom;font-family: SimSun;font-size:10.5pt");
//        //甲乙方
//        Element companyTables = tables.get(0);
//        Elements companyPs = companyTables.getElementsByTag("p");
//        companyPs.attr("style", "line-height:4.0pt;text-autospace:none;vertical-align:\n" +
//                "  bottom;font-family: SimSun;font-size:10.5pt");
        //清除换行
        Element lineP = ps.get(4);
        lineP.remove();
        //供货清单
        Element supplyHead = ps.get(5);
        String supplyHeadStyle = supplyHead.attr("style");
        supplyHeadStyle = supplyHeadStyle.replace("margin-left:47.35pt;", "");
        supplyHead.attr("style", supplyHeadStyle);
        Element supplyTable = tables.get(1);
        Elements supplyRows = supplyTable.getElementsByTag("tr");
        Element supplyHeader = supplyRows.get(0);
//        Elements supplyHeadSpans = supplyHeader.getElementsByTag("span");
//        supplyHeadSpans.attr("style", "font-family: SimSun;font-size:8.5pt;font-weight:bold");
        supplyHeader.html("<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>序号<span lang=EN-US></td>" +
                "<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>品牌<span lang=EN-US></td>" +
                "<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>品名<span lang=EN-US></td>" +
                "<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>型号<span lang=EN-US></td>" +
                "<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>单位<span lang=EN-US></td>" +
                "<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>数量<span lang=EN-US></td>" +
                "<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>单价(元)<span lang=EN-US></td>" +
                "<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>金额(元)<span lang=EN-US></td>" +
                "<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>付款金额(元)<span lang=EN-US></td>" +
                "<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>货期(天)<span lang=EN-US></td>" +
                "<td><span style='font-size:10.0pt; font-family:SimSun; color:black'>备注<span lang=EN-US></td></tr>");

        for (int i = 1; i < supplyRows.size(); i++) {
            Element supplyRow = supplyRows.get(i);
            Elements supplyRowSpans = supplyRow.getElementsByTag("span");
            supplyRowSpans.attr("style", "font-family: SimSun;font-size:8.5pt;color:black'");
        }
        //处理换行
        Elements divBrs = firstDiv.getElementsByTag("br");
        divBrs.get(0).remove();
        divBrs.get(1).remove();
        //处理段落间距
//        for (int i = 5; i < ps.size(); i++) {
//            Element paragraphP = ps.get(i);
//            String paragraphStyle = paragraphP.attr("style");
//            paragraphStyle = paragraphStyle.replace("line-height:20.0pt", "line-height:10.0pt");
//            paragraphP.attr("style", paragraphStyle);
//        }
        //处理段落起始空格
        Element p6 = ps.get(6);
        String p6Style = supplyHead.attr("style");
        p6Style = p6Style.replace("margin-left:10.5pt;", "");
        p6.attr("style", p6Style);
        Element p7 = ps.get(7);
        String p7Style = supplyHead.attr("style");
        p7Style = p7Style.replace("margin-left:21.0pt;", "");
        p7.attr("style", p7Style);
        for (int i = 30; i < ps.size(); i++) {
            if (ps.get(i).html().contains("3") || ps.get(i).html().contains("4")) {
                Element pi = ps.get(i);
                String piStyle = supplyHead.attr("style");
                piStyle = piStyle.replace("margin-left:21.0pt;", "");
                pi.attr("style", piStyle);
            }
        }
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
        List<Element> list = new ArrayList<>();
        for (Element element : select) {
            Elements allElements = element.getAllElements();
            list.addAll(allElements);
        }

        list.stream().forEach(element -> {
            String style = element.attr("style");
            element.attr("style", style + ";font-family: SimSun;font-size:10.5pt");
            if (element.nodeName().equals("p")) {
                String style1 = element.attr("style");
//                style1 = style1.replace("line-height:20.0pt;", "");
                style1 = style1.replace("text-indent:21.1pt;", "");
                style1 = style1.replace("text-indent:25.1pt;", "");
                style1 = style1.replace("text-indent:21.0pt;", "");
                style1 = style1.replace("text-indent:-26.25pt;", "");
                style1 = style1.replace("text-indent:10.5pt;", "");
                element.attr("style", style1);
            }
        });
//        while (list.size() > 0) {
//            Element element = list.get(0);
//            element.attr("style", "font-family: SimSun;font-size:10.5pt");
//            Elements select1 = element.getAllElements();
//            select1.remove(element);
//            list.addAll(select1);
//            list.remove(element);
//        }
    }

//    /**
//     * word转pdf
//     *
//     * @param in             需转换word的输入流
//     * @param targetHtmlPath 生成html的位置
//     * @param targetPDFPath  生成pdf的位置
//     * @param fontPath       字体位置
//     */
//    public static void docToPDF(InputStream in, String targetHtmlPath, String targetXhtmlPath, String targetPDFPath, String fontPath) throws Exception {
//        FileInputStream fileInputStream = null;
//        InputStream inputStream = null;
//        try {
//            HWPFDocument wordDocument = new HWPFDocument(in);
//            Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
//            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document);
//
//            wordToHtmlConverter.processDocument(wordDocument);
//            Document htmlDocument = wordToHtmlConverter.getDocument();
//            DOMSource domSource = new DOMSource(htmlDocument);
//            StreamResult streamResult = new StreamResult(new File(targetHtmlPath));
//            TransformerFactory tf = TransformerFactory.newInstance();
//            Transformer serializer = tf.newTransformer();
//            serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
//            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
//            serializer.setOutputProperty(OutputKeys.METHOD, "html");
//            serializer.transform(domSource, streamResult);
//
//            parseToXhtml(targetHtmlPath, targetXhtmlPath);
//            standardHTML(targetXhtmlPath);
//            convertHtmlToPdf(targetXhtmlPath, targetPDFPath, fontPath);
//        }  finally {
//            try {
//                if (inputStream != null) {
//                    inputStream.close();
//                }
//                if (fileInputStream != null) {
//                    fileInputStream.close();
//                }
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        }
//    }
//
//    public static void docToPDFByTemplate(byte[] in, String targetHtmlPath, String targetXhtmlPath, String targetPDFPath, String fontPath) throws Exception {
//        FileInputStream fileInputStream = null;
//        InputStream inputStream = null;
//        FileOutputStream fileOutputStream = null;
//        try {
//            System.out.println("获取配置文件targetHtmlPath，开始：" + new Date());
//            fileOutputStream = new FileOutputStream(targetHtmlPath);
//            fileOutputStream.write(in);
//            System.out.println("获取配置文件targetHtmlPath，结束：" + new Date());
//            parseToXhtml(targetHtmlPath,targetXhtmlPath);
//            System.out.println("html文件解析成xhtml，变成标准的html文件：" + new Date());
//            standardHTMLByTemplate(targetXhtmlPath);
//            System.out.println("生成模块 xhtml转成标准html文件：" + new Date());
//            convertHtmlToPdf(targetXhtmlPath, targetPDFPath, fontPath);
//            System.out.println("解决中文支持问题："  + new Date());
//        } finally {
//            try {
//                if (inputStream != null) {
//                    inputStream.close();
//                }
//                if (fileInputStream != null) {
//                    fileInputStream.close();
//                }
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        }
//    }
//
//    /*
//     * html文件解析成xhtml，变成标准的html文件
//     */
//    private static void parseToXhtml(String targetPath,String targetXhtmlPath) {
//        ByteArrayOutputStream tidyOutStream = null; // 输出流
//        ByteArrayOutputStream bos = null;
//        ByteArrayInputStream stream = null;
//        FileInputStream fis = null;
//        DataOutputStream to = null;
//        try {
//            fis = new FileInputStream(targetPath);
//            // Reader reader;
//            bos = new ByteArrayOutputStream();
//            int ch;
//            while ((ch = fis.read()) != -1) {
//                bos.write(ch);
//            }
//            byte[] bs = bos.toByteArray();
//            bos.close();
//            String hope_utf_8 = new String(bs, "UTF-8");// 默认是utf-8
//            byte[] hope_b = hope_utf_8.getBytes();
//            String basil = new String(hope_b, "UTF-8");// 在此处可转换成其他字符类型
//            stream = new ByteArrayInputStream(basil.getBytes());
//            tidyOutStream = new ByteArrayOutputStream();
//
//            Tidy tidy = new Tidy();
//            tidy.setInputEncoding("UTF-8");
//            tidy.setQuiet(true);
//            tidy.setOutputEncoding("UTF-8");
//            tidy.setShowWarnings(true);
//            tidy.setIndentContent(true);
//            tidy.setSmartIndent(true);
//            tidy.setIndentAttributes(false);
//            tidy.setWraplen(1024);
//
//            // 输出为xhtml
//            tidy.setXHTML(true);
//            tidy.setErrout(new PrintWriter(System.out));
//            tidy.parse(stream, tidyOutStream);
//            to = new DataOutputStream(new FileOutputStream(targetXhtmlPath));
//            tidyOutStream.writeTo(to);
//        } catch (Exception ex) {
//            System.out.println(ex.toString());
//            ex.printStackTrace();
//        } finally {
//            try {
//                if (to != null) {
//                    to.close();
//                }
//                if (stream != null) {
//                    stream.close();
//                }
//                if (fis != null) {
//                    fis.close();
//                }
//                if (bos != null) {
//                    bos.close();
//                }
//                if (tidyOutStream != null) {
//                    tidyOutStream.close();
//                }
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        }
//    }
//
//    private static boolean convertHtmlToPdf(String inputFile, String outputFile, String fontPath) throws Exception {
//        OutputStream os = new FileOutputStream(outputFile);
//        ITextRenderer renderer = new ITextRenderer();
//        String url = new File(inputFile).toURI().toURL().toString();
//
//        // 解决中文支持问题
//        ITextFontResolver fontResolver = renderer.getFontResolver();
//        fontResolver.addFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
//
//        renderer.setDocument(url);
//        renderer.layout();
//        renderer.createPDF(os);
//        os.flush();
//        os.close();
//        return true;
//    }
//
//    /*
//     * xhtml转成标准html文件
//     * targetHtml:要处理的html文件路径
//     */
//    private static void standardHTML(String targetHtml) {
//        try {
//            File f = new File(targetHtml);
//            org.jsoup.nodes.Document doc = Jsoup.parse(f, "UTF-8");
//
//            doc.select("meta").removeAttr("name");
//            doc.select("meta").attr("content", "text/html; charset=UTF-8");
//            doc.select("meta").attr("http-equiv", "Content-Type");
//            doc.select("style").attr("mce_bogus", "1");
//            doc.select("body").attr("style", "font-family: SimSun");
//            doc.select("html").before("<?xml version='1.0' encoding='UTF-8'>");
//            Elements style = doc.getElementsByTag("style");
//            style.get(0).append("body {\n" +
//                    "            font-family: SimSun;}\n" +
//                    "@page{size: A4}");
//            Elements select = doc.select("body > *");
//            List<Element> list = new ArrayList<>(select);
//            while (list.size() > 0) {
//                Element element = list.get(0);
//                element.attr("style", "font-family: SimSun;");
//                Elements select1 = element.getAllElements();
//                select1.remove(element);
//                list.addAll(select1);
//                list.remove(element);
//            }
//            /*
//             * Jsoup只是解析，不能保存修改，所以要在这里保存修改。
//             */
//            FileOutputStream fos = new FileOutputStream(f, false);
//            OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
//            osw.write(doc.html());
//            osw.close();
//        } catch (UnsupportedEncodingException e) {
//            e.printStackTrace();
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//
//    }
//
//    /*
//     * 生成模块 xhtml转成标准html文件
//     * targetHtml:要处理的html文件路径
//     */
//    public static void standardHTMLByTemplate(String targetHtml) throws IOException {
//        File f = new File(targetHtml);
//        org.jsoup.nodes.Document doc = Jsoup.parse(f, "UTF-8");
//
//        doc.select("meta").removeAttr("name");
//        doc.select("meta").attr("content", "text/html; charset=UTF-8");
//        doc.select("meta").attr("http-equiv", "Content-Type");
//        doc.select("meta").html("&nbsp");
//        doc.select("img").html("&nbsp");
//        doc.select("style").attr("mce_bogus", "1");
//        doc.select("style").html("");
//        doc.select("body").attr("style", "font-family: SimSun");
//        doc.select("html").before("<?xml version='1.0' encoding='UTF-8'>");
//        Elements style = doc.getElementsByTag("style");
//        style.get(0).append("span{" +
//                "            color:black}" +
//                "body {\n" +
//                "            font-family: SimSun;}\n" +
//                "@page { size: A4 }" +
//                "div.header {\n" +
//                "    display: block; text-align: center; \n" +
//                "    position: running(header);\n" +
//                "}\n" +
//                "div.footer {\n" +
//                "    display: block; text-align: right;\n" +
//                "    position: running(footer);\n" +
//                "}\n" +
//                "div.content {page-break-after: always;}\n" +
//                "@page {\n" +
//                "     @top-center { content: element(header) }\n" +
//                "}\n" +
//                "@page { \n" +
//                "    @bottom-center { content: element(footer) }\n" +
//                "}");
//        Elements select = doc.select("body > *");
//        Elements divs = doc.getElementsByTag("div");
//        List<Element> list = new ArrayList<>(select);
//        while (list.size() > 0) {
//            Element element = list.get(0);
//            element.attr("style", "font-family: SimSun;");//;
//            Elements select1 = element.getAllElements();
//            select1.remove(element);
//            list.addAll(select1);
//            list.remove(element);
//        }
//        Element element1 = divs.get(1);
//        element1.attr("align", "left");
//        Elements tables = doc.select("table");
//        if (tables.size() >2) {
//            Element table1 = tables.get(1);
//            table1.attr("border", "1px");
//        }
//        /*2020-12-22 转pdf添加页眉*/
//        Element element3 = divs.get(0);
//        Elements element3ElementsByDiv = element3.getElementsByTag("div");
//        Element element4 = element3ElementsByDiv.get(0);
//        element4.attr("align", "left");
//
//        divs.get(0).before("    <div class='header'><img style=\"height:50pt\" src='http://static.ckmro.com:8082/static/contract.files/image001.png'/></div>\n" +
//                "    <div class='footer'></div>");
//        /*
//         * Jsoup只是解析，不能保存修改，所以要在这里保存修改。
//         */
//        FileOutputStream fos = new FileOutputStream(f, false);
//        OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
//        osw.write(doc.html());
//        osw.close();
//    }
}
