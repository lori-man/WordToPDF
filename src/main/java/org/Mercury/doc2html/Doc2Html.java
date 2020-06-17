package org.Mercury.doc2html;

import fr.opensagres.xdocreport.core.io.internal.ByteArrayOutputStream;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;

public class Doc2Html {

    public static void convert(String sourcePath, String destPath, String imagePath) throws IOException, ParserConfigurationException, TransformerException {
        File sourceFile = new File(sourcePath);
        HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(sourceFile));
        Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document);

        //保存图片，并返回图片路径
        wordToHtmlConverter.setPicturesManager((content, pictureType, suggestedName, widthInches, heightInches) -> {
            try {
                FileOutputStream out = new FileOutputStream(imagePath + "\\" + suggestedName);
                out.write(content);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
            return imagePath + "\\" + suggestedName;
        });
        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();
        DOMSource domSource = new DOMSource(htmlDocument);
        //ByteArrayOutputStream out = new ByteArrayOutputStream();
        org.odftoolkit.odfdom.converter.core.utils.ByteArrayOutputStream out1 = new org.odftoolkit.odfdom.converter.core.utils.ByteArrayOutputStream();
        StreamResult streamResult = new StreamResult(out1);

        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setParameter(OutputKeys.METHOD, "html");
        transformer.transform(domSource, streamResult);
        out1.close();
    }


}
