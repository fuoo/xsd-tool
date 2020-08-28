package com.fuoo.xsd;

import com.fuoo.xsd.utils.WordUtil;

/**
 * @author fuoo
 * @title RunMain
 * @projectName xsd-tool
 * @description TODO
 * @date 2020/8/279:44
 */
public class RunMain {
    public static void main(String[] args) {
        // 存放xsd文档的文件夹地址
        String xsdFileDocument = "C:\\Users\\Administrator\\Desktop\\xsd\\xsd";
        // 接口数字名称 例如:（文件格式名xs---.xsd、xs---a.xsd）文件名为xs1101.xsd、xs1101a.xsd取1101
        String[] implNames = new String[] {"1101"};
        // 结果保存地址
        String resultFile = "C:\\Users\\Administrator\\Desktop\\xsd\\word\\Test.docx";

        try {
            System.out.println("创建表格文档......");
            // 该方法只处理文件名为xs----.xsd或xs----a.xsd的文件
            WordUtil.createXsdDataDocx(xsdFileDocument, resultFile);
            //WordUtil.createXsdDataDocx(xsdFileDocument, resultFile, implNames);
            System.out.println("创建表格文档成功：" + resultFile);
        } catch (Exception e) {
            System.out.println("创建表格文档失败！！！");
            e.printStackTrace();
        }
    }
}
