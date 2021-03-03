# **officeexprot** 使用说明

## 主页

```
https://github.com/kmood/officeexport-java 
```



## 模板语法

```
文本占位符： {key}    例：{title};
图片占位符： {^key^}   例：{^pic^}; 
插入图片需要先插入一张图片，生成xml，然后给图片base64码，后面的那个标签，加一个 alt="{^key^}" 去替换掉原有图片

文本行循环符： [key@alias  key]    例：[titles@t  {t.title}  titles]
单元行循环符： [#key@alias#    #key#]   例：[#titles@t#  {t.title}  #titles#]

ps: 循环符成对出现，会循环输出首位间的所有行。
```



## API

```java
// ActualModelPath是模板文件路径+/，xmlPath是模板文件路径，这两个写一个就可以
DocumentProducer dp = new DocumentProducer(ActualModelPath);
dp.Complie(xmlPath, "filename.xml",true);
dp.produce(map, ExportFilePath);
```



## 快速开始

```java
1、调整word模板，添加占位符，并转换到word 2003 xml文档（.xml）。


2、Maven引入jar包，通过api导出
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