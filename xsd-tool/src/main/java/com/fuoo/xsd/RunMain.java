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
        // xsd文档地址
        String xsdFileDocument = "F:\\xsd";
        // 接口数字名称 例如:（文件格式名xs---.xsd、xs---a.xsd）文件名为xs1101.xsd、xs1101a.xsd取1101
        String[] implNames = new String[] {"1101"};
        // 结果保存地址
        String resultFile = "F:\\docx\\Test.docx";

        try {
            System.out.println("创建中......");
            WordUtil.createXsdDataDocx(xsdFileDocument, resultFile, implNames);
            System.out.println("创建成功：" + resultFile);
        } catch (Exception e) {
            System.out.println("创建失败！！！");
            e.printStackTrace();
        }

    }
}
