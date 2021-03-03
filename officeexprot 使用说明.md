# **officeexprot** 使用说明

<https://github.com/kmood/officeexport-java> 

## 简述

**officeexprot** 是一个极简的word模板动态导出库。 基于Apache FreeMarker，遵从模板 + 数据模型 = 输出的理念， 通过极简API实现javaBean即数据源，模板即样式的Word、Excel导出，主要实现以下功能：

- 基本文本的输出，文本占位符样式即输出文本样式。
- 文本行、表格行单行或多行的遍历输出，并能够进行循环嵌套。
- 提供数据处理的插件，通过添加处理器可定制任意输出值，例如：特定项的日期、数字等文本格式转换
- 图片保留样式的输出。

## 基本入门

## 模板语法

  **文本占位符：** {key}    例：{title};

  **图片占位符：** {^key^}   例：{^pic^};

  **文本行循环符：** [*key@alias*  *key*]    例：[*titles@t*  {t.title}  *titles*]

  **单元行循环符：** [#key@alias#    #key#]   例：[#titles@t#  {t.title}  #titles#]

  **ps:** 循环符成对出现，会循环输出首位间的所有行。

## API

```java
DocumentProducer dp = new DocumentProducer(ActualModelPath);
dp.Complie(xmlPath, "filename.xml",true);
dp.produce(map, ExportFilePath);
```



## 快速开始

> > 1、调整word模板，添加占位符，并转换到word 2003 xml文档（.xml）。

> > 2、Maven引入jar包，通过api导出

```java
   <dependency>
       <groupId>com.github.kmood</groupId>
       <artifactId>officeexport-java</artifactId>
       <version>1.0.1.3-RELEASE</version>
   </dependency>
   
   
  HashMap<String, Object> data = new HashMap<>();
  ...准备数据
  data.put("zxsm",zxsmList);
  data.put("sbsm","kmood-导出-商标说明");
  
  DocumentProducer dp = new DocumentProducer(ActualGenerateModelPath);
  dp.Complie(xmlFilePath, "2003xml文件名.xml",true);
  dp.produce(map, ExportFile);
```