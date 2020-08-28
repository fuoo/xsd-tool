package com.fuoo.xsd.utils;

import com.fuoo.xsd.constants.XMLConstants;
import com.fuoo.xsd.entity.XSDNode;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.SAXReader;

import java.io.File;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author fuoo
 * @title Xsd解析工具
 * @projectName xsd-tool
 * @description TODO
 * @date 2020/8/2611:24
 */
public class ParseXsdUtil {
    private static List<XSDNode> list = new ArrayList<>();

    /**
     * 解析XSD，返回数据节点对象列表
     * @param filePath
     * @return
     * @throws Exception
     */
    public static List<XSDNode> paserXSD(File filePath, String dateType) throws Exception {
        list.clear();

        SAXReader saxReader = new SAXReader();

        Document doc = saxReader.read(filePath);

        Element element = doc.getRootElement();

        String basePath = null;
        Element dataElement = null;
        if("".equals(XMLConstants.XSD_DEFAULT_NAMESPACE)){
            if("".equals(dateType)){
                dataElement = element;
            }else{
                basePath = "//element[@name=\"" + dateType + "\"]";
                dataElement = (Element) element.selectSingleNode(basePath);
            }
        }else{
            basePath = "//" + XMLConstants.XSD_DEFAULT_NAMESPACE + ":complexType[@name=\"" + dateType + "\"]";
            dataElement = (Element) element.selectSingleNode(basePath);
        }

        String elementPath = null;
        if("".equals(XMLConstants.XSD_DEFAULT_NAMESPACE)){
            elementPath = "//element";
        }else{
            elementPath = "//" + XMLConstants.XSD_DEFAULT_NAMESPACE + ":complexType";
        }

        paseData(dateType, dataElement, "//", elementPath, "//");

        return list;
    }

    /**
     * 转换XSD的数据节点，生成XSDNode对象
     *
     * @param element
     * @param xPath
     * @param xsdPath
     * @param unboundedXpath
     */
    private static void paseData(String dateType, Element element, String xPath, String xsdPath, String unboundedXpath) {

        if(element==null) {
            return;
        }

        // 获取节点name属性
        String nodeName = element.attributeValue("name");

        // 组装xml文档中节点的XPath

        xPath += nodeName;

        unboundedXpath += nodeName;

        // 并列多节点限制属性
        String maxOccurs = element.attributeValue("maxOccurs");

        if (maxOccurs != null && !"1".equals(maxOccurs) && !("//" + dateType + "").equals(xPath)) {// 节点可以有多个

            unboundedXpath += XMLConstants.XSD_UNBOUNDED;

        }

        // 组装下一个element元素的XPath

        String currentXsdPath = xsdPath + "[@name=\"" + nodeName + "\"]" + "/" + XMLConstants.XSD_DEFAULT_NAMESPACE

                + ":sequence/" + XMLConstants.XSD_DEFAULT_NAMESPACE

                + ":element";

        // 查找该节点下所有的element元素

        List<Node> elementNodes = element.selectNodes(currentXsdPath);

        if (elementNodes != null && elementNodes.size() > 0) {// 如果下面还有element,说明不是叶子

            Iterator<Node> nodes = elementNodes.iterator();

            while (nodes.hasNext()) {

                if (!xPath.endsWith("/")) {

                    xPath += "/";

                    unboundedXpath += "/";

                }

                Element ele = (Element) nodes.next();

                paseData(dateType, ele, xPath, currentXsdPath, unboundedXpath);

            }

        } else { // 该element为叶子

            XSDNode xsdNode = new XSDNode();

            // 获取注释节点

            String annotation = "";

            Node annotationNode = element

                    .selectSingleNode(xsdPath + "[@name=\"" + nodeName + "\"]/" + XMLConstants.XSD_DEFAULT_NAMESPACE

                            + ":annotation/" + XMLConstants.XSD_DEFAULT_NAMESPACE + ":documentation");

            if (annotationNode != null) {
                annotation = annotationNode.getText();
            }

            // 获取节点类型属性
            String nodeType = "";

            Attribute type = element.attribute("type");

            if (type != null){

                nodeType = type.getText();

            }else {

                String spath = xsdPath + "[@name=\"" + nodeName + "\"]/" + XMLConstants.XSD_DEFAULT_NAMESPACE + ":simpleType/"

                        + XMLConstants.XSD_DEFAULT_NAMESPACE + ":restriction";

                Element typeNode = (Element) element.selectSingleNode(spath);

                if (typeNode != null) {

                    Attribute base = typeNode.attribute("base");

                    if (base != null)

                        nodeType = base.getText();

                }

            }

            xsdNode.setMaxOccurs(element.attribute("maxOccurs") == null ? null:element.attribute("maxOccurs").getText());
            xsdNode.setMinOccurs(element.attribute("minOccurs") == null ? null:element.attribute("minOccurs").getText());
            xsdNode.setName(nodeName);
            xsdNode.setXPath(xPath);
            xsdNode.setAnnotation(annotation);
            xsdNode.setType(nodeType.replaceAll("xs:", ""));
            xsdNode.setUnboundedXpath(unboundedXpath);
            list.add(xsdNode);
        }

    }
}