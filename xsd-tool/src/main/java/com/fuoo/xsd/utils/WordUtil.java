package com.fuoo.xsd.utils;

import com.fuoo.xsd.constants.XMLConstants;
import com.fuoo.xsd.entity.XSDNode;
import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import java.awt.*;
import java.io.*;
import java.util.*;
import java.util.List;


/**
 * @author fuoo
 * @title word文档生成
 * @projectName xsd-tool
 * @description TODO
 * @date 2020/8/2616:36
 */
public class WordUtil {
    // 文件前后缀
    public final static String FILE_PREFIX = "xs";
    public final static String FILE_SUFFIX = ".xsd";
    public final static String FILE_SUFFIX_2 = "a.xsd";
    public final static String DOCX_SUFFIX = ".docx";
    // 临时文件地址
    public final static File FILE_WORD_DOCUMENT = new File("D:\\xsd_tool");

    static {
        if (!FILE_WORD_DOCUMENT.exists()) {
            FILE_WORD_DOCUMENT.mkdirs();
        }
    }

    /**
    　* @description 创建xsd文件数据的docx文档
      * 该方法只处理文件名为xs----.xsd或xs----a.xsd的文件
      *@param: [xsdFileDocument xsd文件夹, resultFile 结果文件]
    　* @return: void
    　* @author: fuoo
    　* @date: 2020/8/28 10:36
    　*/
    public static void createXsdDataDocx(String xsdFileDocument, String resultFile) throws Exception{
        File xsdFileDocumentFile = new File(xsdFileDocument);
        List<File> xsdFileList = Arrays.asList(xsdFileDocumentFile.listFiles());
        for (File file:xsdFileList) {
            String fileName = file.getName();
            if (!fileName.matches("xs[0-9]{4}(a.xsd|.xsd)")) {
                xsdFileList.remove(file);
            }
        }
        createXsdDataDocx(xsdFileList, resultFile);
    }

    /**
    　* @description: 创建xsd文件数据的docx文档
      *该方法只处理文件名为xs----.xsd或xs----a.xsd的文件
    　* @param: [xsdFileDocument xsd文件夹, resultFile 结果文件, implNames接口名称]
    　* @return: void
    　* @author: fuoo
    　* @date: 2020/8/27 15:44
    　*/
    public static void createXsdDataDocx(String xsdFileDocument, String resultFile, String[] implNames) throws Exception{
        List<File> xsdFileList = WordUtil.readXsdFile(xsdFileDocument, implNames);
        createXsdDataDocx(xsdFileList, resultFile);
    }

    private static void createXsdDataDocx(List<File> xsdFileList, String resultFile) throws Exception{
        for (File file:xsdFileList) {
            if (file.getName().contains(WordUtil.FILE_SUFFIX_2)) {
                // 请响应文件
                List<XSDNode> xsdNodeList = ParseXsdUtil.paserXSD(file, XMLConstants.PAGE_INFO);
                String fileName = WordUtil.createFileName(file, XMLConstants.PAGE_INFO);
                WordUtil.createWordTable(fileName, xsdNodeList);

                // accessoryContents的XML数据内容
                xsdNodeList = ParseXsdUtil.paserXSD(file, XMLConstants.ACCESSORY_CONTENTS_TYPE);
                fileName =  WordUtil.createFileName(file, XMLConstants.ACCESSORY_CONTENTS_TYPE);
                WordUtil.createWordTable(fileName, xsdNodeList);
            } else {
                // 请求体
                List<XSDNode> xsdNodeList = ParseXsdUtil.paserXSD(file, XMLConstants.REQUEST_DETAIL);
                String fileName = WordUtil.createFileName(file, XMLConstants.REQUEST_DETAIL);
                WordUtil.createWordTable(fileName, xsdNodeList);
            }
            System.out.println(file + "生成表格成功！");
        }

        // 文件合并
        File[] listFiles = WordUtil.FILE_WORD_DOCUMENT.listFiles();
        WordUtil.mergeWordList(resultFile, Arrays.asList(listFiles));

        // 删除临时文件
        for (File file:listFiles) {
            file.delete();
        }
        WordUtil.FILE_WORD_DOCUMENT.delete();
    }


    /**
    　* @description: 创建word表格。
    　* @param: filePath
      * @param: xsdNodeList
    　* @return: void
    　* @author: fuoo
    　* @date: 2020/8/27 9:41
    　*/
    public static void createWordTable(String filePath, List<XSDNode> xsdNodeList) {
        //创建Document对象
        Document doc = new Document();
        Section sec = doc.addSection();

        //======================标题======================
        String fileName =  new File(filePath).getName().replaceAll("0_|1_|2_|.docx", "");
        String implName = fileName.replaceAll("[^0-9]", "");
        if (fileName.contains(XMLConstants.REQUEST_DETAIL)) {
            Paragraph paragraph = sec.addParagraph();
            TextRange textRange = paragraph.appendText("方法参数：BusinessType，取值为“" + implName +"”。\n" +
                    "方法参数：RequestContent，XML消息体数据格式定义如下：");
            textRange.getCharacterFormat().setFontSize(10.5f);
        } else if (fileName.contains(XMLConstants.PAGE_INFO)) {
            Paragraph paragraph = sec.addParagraph();
            TextRange textRange = paragraph.appendText("返回结果：PageInfo消息体数据格式定义如下：" + implName);
            textRange.getCharacterFormat().setFontSize(10.5f);
        } else if (fileName.contains(XMLConstants.ACCESSORY_CONTENTS_TYPE)) {
            Paragraph paragraph = sec.addParagraph();
            TextRange textRange = paragraph.appendText("AccessoryContents返回内容" + implName);
            textRange.getCharacterFormat().setFontSize(10.5f);
        }

        if (xsdNodeList == null || xsdNodeList.size() == 0) {
            Paragraph paragraph = sec.addParagraph();
            paragraph.appendText("无数据!!!");
            return;
        }

        //======================表格生成======================
        //声明数组内容
        String[] header = {"名称","数据类型","是否可空", "说明"};
        String[][] data = handleXSDNodeData(xsdNodeList);

        //添加表格
        Table table = sec.addTable(true);
        // 设置列宽
        float b = 28.3286119f;
        table.setColumnWidth(new float[]{3.98f*b, 2.24f*b, 2.15f*b, 6.47f*b});
        //设置行数和列数
        table.resetCells(data.length + 1, header.length);
        // 表格样式
        //table.applyStyle(DefaultTableStyle.Colorful_List);
        // 对齐方式 居中
        table.getTableFormat().setHorizontalAlignment(RowAlignment.Center);
        // 设置表格第一行作为表头，写入表头数组内容，并格式化表头数据
        TableRow row = table.getRows().get(0);
        row.isHeader(true);
        row.getRowFormat().setBackColor(new Color(192,192,192));
        row.setHeightType(TableRowHeightType.Auto);
        //row.setHeight(b);
        for (int i = 0; i < header.length; i++) {
            row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
            Paragraph p = row.getCells().get(i).addParagraph();
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
            TextRange range1 = p.appendText(header[i]);
            range1.getCharacterFormat().setFontName("宋体");
            range1.getCharacterFormat().setFontSize(12f);
            range1.getCharacterFormat().setBold(true);
            range1.getCharacterFormat().setTextColor(Color.BLACK);
            // 设置行距
            p.getFormat().setBeforeAutoSpacing(false);
            p.getFormat().setBeforeSpacing(3f);
            p.getFormat().setAfterAutoSpacing(false);
            p.getFormat().setAfterSpacing(3f);
        }

        //写入剩余组内容到表格，并格式化数据
        for (int r = 0; r < data.length; r++) {
            TableRow dataRow = table.getRows().get(r + 1);
            dataRow.setHeightType(TableRowHeightType.At_Least);
            dataRow.getRowFormat().setBackColor(Color.white);
            for (int c = 0; c < data[r].length; c++) {
                dataRow.getCells().get(c).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                TextRange range2 = dataRow.getCells().get(c).addParagraph().appendText(data[r][c]);
                range2.getOwnerParagraph().getFormat().setHorizontalAlignment(HorizontalAlignment.Justify);
                range2.getCharacterFormat().setFontName("Calibri");
                range2.getCharacterFormat().setFontSize(10.5f);
            }
        }
        //设置表格边框样式
        //table.getTableFormat().getBorders().setBorderType(BorderStyle.Thick_Thin_Large_Gap);
        //保存文档
        doc.saveToFile(filePath, FileFormat.Docx);
    }

    /**
    　* @description: xsd节点数据转换
    　* @param: [xsdNodeList]
    　* @return: java.lang.String[][]
    　* @author: fuoo
    　* @date: 2020/8/27 9:40
    　*/
    private static String[][] handleXSDNodeData(List<XSDNode> xsdNodeList) {
        if (xsdNodeList.size() == 0) {
            return null;
        }

        String[][] data = new String[xsdNodeList.size()][4];
        int i = 0;
        for (XSDNode node : xsdNodeList) {
            if ("0".equals(node.getMinOccurs()) && "1".equals(node.getMaxOccurs())) {
                node.setIsCanNull("是");
            } else {
                node.setIsCanNull("否");
            }

            if (!"int".equals(node.getType())) {
                char[] str = node.getType().toCharArray();
                if (str.length >= 1 && str[0] >= 'a' && str[0] <= 'z') {
                    str[0] -= 32;
                }
                node.setType(new String(str));
            }
            data[i] = new String[]{node.getName(), node.getType(), node.getIsCanNull(), ""};
            i++;
        }

        return data;
    }

    /**
    　* @description: 读取文件
    　* @param: [parentPath, imptName]
    　* @return: java.io.File[]
    　* @author: fuoo
    　* @date: 2020/8/27 12:14
    　*/
    public static List<File> readXsdFile(String parentPath, String... implNames) {
        File parentFile = new File(parentPath);
        if (!(parentFile.exists() && parentFile.isDirectory())) {
            throw new RuntimeException("文件夹路径错误");
        }

        List<File> fileList = new ArrayList<>();
        List<File> notFoundList = new ArrayList<>();
        for (String imptName : implNames) {
            File file1 = new File(parentPath + File.separator + FILE_PREFIX + imptName + FILE_SUFFIX);
            File file2 = new File(parentPath + File.separator + FILE_PREFIX + imptName + FILE_SUFFIX_2);

            if (file1.exists() && file1.isFile()) {
                fileList.add(file1);
            } else {
                notFoundList.add(file1);
            }

            if (file2.exists() && file2.isFile()) {
                fileList.add(file2);
            } else {
                notFoundList.add(file2);
            }
        }

        if (fileList.size() == 0) {
            throw new RuntimeException("一个对应的文件都未找到");
        }

        if (notFoundList.size() > 0) {
            System.out.println("未找到文件：");
            for (File file:notFoundList) {
                System.out.println(file.getPath());
            }
        }
        return fileList;
    }

    /**
    　* @description: 创建临时文件
    　* @param: [sourceFile, type]
    　* @return: java.lang.String
    　* @author: fuoo
    　* @date: 2020/8/27 15:22
    　*/
    public static String createFileName(File sourceFile, String type) {
        String str = "";
        switch (type) {
            case XMLConstants.REQUEST_DETAIL:
                str = "0_";
                break;
            case XMLConstants.PAGE_INFO:
                str = "1_";
                break;
            case XMLConstants.ACCESSORY_CONTENTS_TYPE:
                str = "2_";
                break;
            default:
                str = "";
        }
        return FILE_WORD_DOCUMENT.getPath() + File.separator + str + sourceFile.getName().replaceAll("[^0-9]", "") + type + WordUtil.DOCX_SUFFIX;
    }

    /**
     　* @description: 文件合成
     　* @param: [saveFilePath, srcfile]
     　* @return: void
     　* @author: fuoo
     　* @date: 2020/8/27 11:31
     　*/
    public static void mergeWordList(String saveFilePath, List<File> srcfile) throws Exception {
        File newFile = new File(saveFilePath);
        newFile.delete();
        OutputStream dest = new FileOutputStream(newFile);
        ArrayList<XWPFDocument> documentList = new ArrayList<>();
        XWPFDocument doc = null;
        for (int i = 0; i < srcfile.size(); i++) {
            FileInputStream in = new FileInputStream(srcfile.get(i).getPath());
            OPCPackage open = OPCPackage.open(in);
            XWPFDocument document = new XWPFDocument(open);
            documentList.add(document);
        }
        for (int i = 0; i < documentList.size(); i++) {
            doc = documentList.get(0);
            if(i == 0){
                //首页直接分页，不再插入首页文档内容
                //documentList.get(i).createParagraph().createRun().addBreak(org.apache.poi.xwpf.usermodel.BreakType.PAGE);
                //appendBody(doc,documentList.get(i));
            }else if(i == documentList.size()-1){
                //尾页不再分页，直接插入最后文档内容
                appendBody(doc,documentList.get(i));
            }else{
                //documentList.get(i).createParagraph().createRun().addBreak(org.apache.poi.xwpf.usermodel.BreakType.PAGE);
                appendBody(doc,documentList.get(i));
            }
        }
        doc.write(dest);
    }

    private static void appendBody(XWPFDocument src, XWPFDocument append) throws Exception {
        CTBody src1Body = src.getDocument().getBody();
        CTBody src2Body = append.getDocument().getBody();

        List<XWPFPictureData> allPictures = append.getAllPictures();
        // 记录图片合并前及合并后的ID
        Map<String,String> map = new HashMap<String,String>();
        for (XWPFPictureData picture : allPictures) {
            String before = append.getRelationId(picture);
            //将原文档中的图片加入到目标文档中
            String after = src.addPictureData(picture.getData(), org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_PNG);
            map.put(before, after);
        }
        appendBody(src1Body, src2Body,map);
    }

    private static void appendBody(CTBody src, CTBody append,Map<String,String> map) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = append.xmlText(optionsOuter);

        String srcString = src.xmlText();
        String prefix = srcString.substring(0,srcString.indexOf(">")+1);
        String mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
        String sufix = srcString.substring( srcString.lastIndexOf("<") );
        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        if (map != null && !map.isEmpty()) {
            //对xml字符串中图片ID进行替换
            for (Map.Entry<String, String> set : map.entrySet()) {
                addPart = addPart.replace(set.getKey(), set.getValue());
            }
        }
        //将两个文档的xml内容进行拼接
        CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
        src.set(makeBody);
    }
}