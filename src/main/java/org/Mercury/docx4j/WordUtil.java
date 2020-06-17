package org.Mercury.docx4j;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.docx4j.convert.out.pdf.PdfConversion;
import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTAltChunk;

/**
 * @author HMZ
 */
public class WordUtil {
    public void mergeDocx(HttpServletResponse response,List<String> list){
        List<InputStream>  inList=new ArrayList<InputStream>();
        for(int i=0;i<list.size();i++)
            try {
                inList.add(new FileInputStream(list.get(i)));
            } catch (FileNotFoundException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        try {
            InputStream inputStream=mergeDocx(inList);
            saveTemplate(response,inputStream);
        } catch ( IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (Docx4JException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    public InputStream mergeDocx(final List<InputStream> streams)  throws Docx4JException, IOException {

        WordprocessingMLPackage target = null;
        final File generated = File.createTempFile("generated", ".docx");

        int chunkId = 0;
        Iterator<InputStream> it = streams.iterator();
        while (it.hasNext()) {
            InputStream is = it.next();
            if (is != null) {
                if (target == null) {
                    // Copy first (master) document
                    OutputStream os = new FileOutputStream(generated);
                    os.write(IOUtils.toByteArray(is));
                    os.close();

                    target = WordprocessingMLPackage.load(generated);
                } else {
                    // Attach the others (Alternative input parts)
                    insertDocx(target.getMainDocumentPart(),
                            IOUtils.toByteArray(is), chunkId++);
                }
            }
        }

        if (target != null) {
            target.save(generated);
            FileInputStream input = new FileInputStream(generated);
            //获取docx解析对象
//            XWPFDocument document = new XWPFDocument(target);
            return new FileInputStream(generated);
        } else {
            return null;
        }
    }

    // 插入文档
    private void insertDocx(MainDocumentPart main, byte[] bytes, int chunkId) {
        try {
            AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(
                    new PartName("/part" + chunkId + ".docx"));
            //   afiPart.setContentType(new ContentType(CONTENT_TYPE));
            afiPart.setBinaryData(bytes);
            Relationship altChunkRel = main.addTargetPart(afiPart);

            CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();
            chunk.setId(altChunkRel.getId());

            main.addObject(chunk);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public  void saveTemplate(HttpServletResponse response,InputStream fis) throws IOException{
        OutputStream fos = response.getOutputStream();
        int bytesum = 0;
        int byteread = 0;
        try {
            byte[] buffer = new byte[10240];
            while ( (byteread = fis.read(buffer)) != -1) {
                bytesum += byteread; //字节数 文件大小
                fos.write(buffer, 0, byteread);
            }
            fis.close();
            fos.close();
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * docx文档转换为PDF
     * @param docx docx文档
     * @param pdfPath PDF文档存储路径
     * @throws Exception 可能为Docx4JException, FileNotFoundException, IOException等
     */
    @SuppressWarnings("deprecation")
    public void convertDocxToPDF(File docx, String pdfPath) throws Exception {
        OutputStream os = null;
        try {
            WordprocessingMLPackage mlPackage = WordprocessingMLPackage.load(docx);
//            Mapper fontMapper = new BestMatchingMapper();
            Mapper fontMapper = new IdentityPlusMapper();
            //中文字体转换
            fontMapper.getFontMappings().put("华文行楷", PhysicalFonts.getPhysicalFonts().get("STXingkai"));
            fontMapper.getFontMappings().put("隶书", PhysicalFonts.getPhysicalFonts().get("LiSu"));
            fontMapper.getFontMappings().put("宋体",PhysicalFonts.getPhysicalFonts().get("SimSun"));
            fontMapper.getFontMappings().put("微软雅黑",PhysicalFonts.getPhysicalFonts().get("Microsoft Yahei"));
            fontMapper.getFontMappings().put("黑体",PhysicalFonts.getPhysicalFonts().get("SimHei"));
            fontMapper.getFontMappings().put("楷体",PhysicalFonts.getPhysicalFonts().get("KaiTi"));
            fontMapper.getFontMappings().put("新宋体",PhysicalFonts.getPhysicalFonts().get("NSimSun"));
            fontMapper.getFontMappings().put("华文行楷", PhysicalFonts.getPhysicalFonts().get("STXingkai"));
            fontMapper.getFontMappings().put("华文仿宋", PhysicalFonts.getPhysicalFonts().get("STFangsong"));
            fontMapper.getFontMappings().put("宋体扩展",PhysicalFonts.getPhysicalFonts().get("simsun-extB"));
            fontMapper.getFontMappings().put("仿宋",PhysicalFonts.getPhysicalFonts().get("FangSong"));
            fontMapper.getFontMappings().put("仿宋_GB2312",PhysicalFonts.getPhysicalFonts().get("FangSong_GB2312"));
            fontMapper.getFontMappings().put("幼圆",PhysicalFonts.getPhysicalFonts().get("YouYuan"));
            fontMapper.getFontMappings().put("华文宋体",PhysicalFonts.getPhysicalFonts().get("STSong"));
            fontMapper.getFontMappings().put("华文中宋",PhysicalFonts.getPhysicalFonts().get("STZhongsong"));
            mlPackage.setFontMapper(fontMapper);

            PdfConversion conversion = new org.docx4j.convert.out.pdf.viaXSLFO.Conversion(mlPackage);
            os = new FileOutputStream(pdfPath);

            conversion.output(os, new PdfSettings());
        } finally {
            IOUtils.closeQuietly(os);
        }
    }

    /**
     * 将文件内容写入到临时文件中
     * @param inputUrl
     * @param file
     * @throws IOException
     */
    public void createFile(String inputUrl,File file) throws IOException{
        //获取docx解析对象
        XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(inputUrl));
        FileOutputStream stream = new FileOutputStream(file);
        document.write(stream);
        stream.close();
    }

    /**
     * 替换文本
     * @param document docx解析对象
     */
    public static void changeText(XWPFDocument document, String oldValue, String newValue){
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落时候需要进行替换
            String text = paragraph.getText();
            if(text.indexOf("^m^p")!= -1 || text.indexOf("^p^m")!= -1){
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //替换模板原来位置
                    run.setText("^m",0);
                }
            }
        }
    }

}
