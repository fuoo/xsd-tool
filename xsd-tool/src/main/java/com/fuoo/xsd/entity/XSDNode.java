package com.fuoo.xsd.entity;

import lombok.Data;

/**
 * @author fuoo
 * @title n
 * @projectName xsd-tool
 * @description TODO
 * @date 2020/8/2611:22
 */
@Data
public class XSDNode {

    // 节点名称
    private String name;

    // 节点XPath
    private String xPath;

    // 节点描述
    private String annotation;

    // 节点类型
    private String type;

    // 业务用路径,描述路径中的unbound节点
    private String unboundedXpath;

    private String isUnbounded;
    // 是否能为空
    private String isCanNull;

    private String maxOccurs;

    private String minOccurs;
}
