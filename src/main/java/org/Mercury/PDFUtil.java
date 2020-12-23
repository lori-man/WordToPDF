package org.Mercury;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.ExceptionConverter;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;

public class PDFUtil {

    //页码事件
    private static class PageXofYTest extends PdfPageEventHelper{
        public PdfTemplate total;

        public BaseFont bfChinese;

        /**
         * 重写PdfPageEventHelper中的onOpenDocument方法
         */
        @Override
        public void onOpenDocument(PdfWriter writer, Document document) {
            // 得到文档的内容并为该内容新建一个模板
            total = writer.getDirectContent().createTemplate(500, 500);
            try {

                String prefixFont = "";
                String os = System.getProperties().getProperty("os.name");
                if(os.startsWith("win") || os.startsWith("Win")){
                    prefixFont = "C:\\Windows\\Fonts" + File.separator;
                }else {
                    prefixFont = "/usr/share/fonts/chinese" + File.separator;
                }

                // 设置字体对象为Windows系统默认的字体
                bfChinese = BaseFont.createFont(prefixFont + "simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            } catch (Exception e) {
                throw new ExceptionConverter(e);
            }
        }

        /**
         * 重写PdfPageEventHelper中的onEndPage方法
         */
        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            // 新建获得用户页面文本和图片内容位置的对象
            PdfContentByte pdfContentByte = writer.getDirectContent();
            // 保存图形状态
            pdfContentByte.saveState();
            String text = writer.getPageNumber() + "/";
            // 获取点字符串的宽度
            float textSize = bfChinese.getWidthPoint(text, 9);
            pdfContentByte.beginText();
            // 设置随后的文本内容写作的字体和字号
            pdfContentByte.setFontAndSize(bfChinese, 9);

            // 定位'X/'
            float x = (document.right() + document.left()) / 2;
            float y = 56f;
            pdfContentByte.setTextMatrix(x, y);
            pdfContentByte.showText(text);
            pdfContentByte.endText();

            // 将模板加入到内容（content）中- // 定位'Y'
            pdfContentByte.addTemplate(total, x + textSize, y);

            pdfContentByte.restoreState();
        }

        /**
         * 重写PdfPageEventHelper中的onCloseDocument方法
         */
        @Override
        public void onCloseDocument(PdfWriter writer, Document document) {
            total.beginText();
            try {
                String prefixFont = "";
                String os = System.getProperties().getProperty("os.name");
                if(os.startsWith("win") || os.startsWith("Win")){
                    prefixFont = "C:\\Windows\\Fonts" + File.separator;
                }else {
                    prefixFont = "/usr/share/fonts/chinese" + File.separator;
                }

                bfChinese = BaseFont.createFont(prefixFont + "simsun.ttc,0",BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                total.setFontAndSize(bfChinese, 9);
            } catch (DocumentException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
            total.setTextMatrix(0, 0);
            // 设置总页数的值到模板上，并应用到每个界面
            total.showText(String.valueOf(writer.getPageNumber() - 1));
            total.endText();
        }
    }

    //页眉事件
    private static class Header extends PdfPageEventHelper {
        public static PdfPTable header;

        public Header(PdfPTable header) {
            Header.header = header;
        }

        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            //把页眉表格定位
            header.writeSelectedRows(0, -1, 36, 806, writer.getDirectContent());
        }

        /**
         * 设置页眉
         * @param writer
         * @param req
         * @throws MalformedURLException
         * @throws IOException
         * @throws DocumentException
         */
        public void setTableHeader(PdfWriter writer) throws MalformedURLException, IOException, DocumentException {
            String imageAddress = "E://TESTPDF/";
            PdfPTable table = new PdfPTable(1);
            table.setTotalWidth(555);
            PdfPCell cell = new PdfPCell();
            cell.setBorder(0);
            Image image01;
            image01 = Image.getInstance(imageAddress + "testhead.png"); //图片自己传
            //image01.scaleAbsolute(355f, 10f);
            image01.setWidthPercentage(80);
            cell.setPaddingLeft(30f);
            cell.setPaddingTop(-20f);
            cell.addElement(image01);
            table.addCell(cell);
            Header event = new Header(table);
            writer.setPageEvent(event);
        }
    }

    //页脚事件
    private static class Footer extends PdfPageEventHelper {
        public static PdfPTable footer;

        @SuppressWarnings("static-access")
        public Footer(PdfPTable footer) {
            this.footer = footer;
        }

        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            //把页脚表格定位
            footer.writeSelectedRows(0, -1, 38, 50, writer.getDirectContent());
        }

        /**
         * 页脚是图片
         * @param writer
         * @throws MalformedURLException
         * @throws IOException
         * @throws DocumentException
         */
        public void setTableFooter(PdfWriter writer) throws MalformedURLException, IOException, DocumentException {
            String imageAddress = "E://TESTPDF/";
            PdfPTable table = new PdfPTable(1);
            table.setTotalWidth(523);
            PdfPCell cell = new PdfPCell();
            cell.setBorder(1);
            Image image01;
            image01 = Image.getInstance(imageAddress + "testfooter.png"); //图片自己传
            image01.scaleAbsoluteWidth(523);
            image01.scaleAbsoluteHeight(30f);
            image01.setWidthPercentage(100);
            cell.addElement(image01);
            table.addCell(cell);
            Footer event = new Footer(table);
            writer.setPageEvent(event);
        }

        /**
         * 页脚是文字
         * @param writer
         * @param songti09
         */
        public void setTableFooter(PdfWriter writer, Font songti09) {
            PdfPTable table = new PdfPTable(1);
            table.setTotalWidth(520f);
            PdfPCell cell = new PdfPCell();
            cell.setBorder(1);
            String string = "地址:  XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX         网址:  www.xxxxxxx.com       咨询热线:  400x-xxx-xxx";
            Paragraph p = new Paragraph(string, songti09);
            cell.setPaddingLeft(10f);
            cell.setPaddingTop(-2f);
            cell.addElement(p);
            table.addCell(cell);
            Footer event = new Footer(table);
            writer.setPageEvent(event);
        }
    }

    public static void main(String[] args) throws Exception {
        Document document = new Document(PageSize.A4, 48, 48, 60, 65);
        // add index page.
        String path = "test.pdf";
        String dir = "E://TEST";
        File file = new File(dir);
        if (!file.exists()) {
            file.mkdir();
        }
        path = dir + File.separator + path;
        FileOutputStream os = new FileOutputStream(path);
        PdfWriter writer = PdfWriter.getInstance(document, os);

        // 设置页面布局
        writer.setViewerPreferences(PdfWriter.PageLayoutOneColumn);
        // 为这篇文档设置页面事件(X/Y)
        writer.setPageEvent(new PageXofYTest());

        String prefixFont = "";
        String oss = System.getProperties().getProperty("os.name");
        if(oss.startsWith("win") || oss.startsWith("Win")){
            prefixFont = "C:\\Windows\\Fonts" + File.separator;
        }else {
            prefixFont = "/usr/share/fonts/chinese" + File.separator;
        }
        BaseFont baseFont1 = BaseFont.createFont(prefixFont + "simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        Font songti09 = new Font(baseFont1, 9f); //宋体 小五

        document.open();

        //document.newPage();
        PdfPTable pdfPTable = new PdfPTable(1);
        // 为报告添加页眉，事件的发生是在生成报告之后，写入到硬盘之前
        //Header headerTable = new Header(pdfPTable);
        //headerTable.setTableHeader(writer);
        //Footer footerTable = new Footer(pdfPTable);
        //footerTable.setTableFooter(writer, songti09);
        document.add(pdfPTable);

        for (int i = 0; i < 80; i++) {
            document.add(new Paragraph("the first page"));
        }

        //document.newPage();
        document.add(new Paragraph("the second page"));

        document.close();
        os.close();
    }
}
