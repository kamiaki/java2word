package com.aki.java2word.controller;

import com.aki.java2word.po.Study;
import com.aki.java2word.po.UserInfo;
import com.aki.java2word.util.ResultBean;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import org.springframework.stereotype.Controller;
import org.springframework.util.Base64Utils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author: GaoLei
 * @Date: 2019/10/16 11:59
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description: Excel ,word 导出
 */
@Controller
public class BaseController {
    /**
     * 导出用户 Word
     * @throws IOException
     * @return
     */
    @RequestMapping("user/doc")
    @ResponseBody
    public ResultBean exportDoc() throws  IOException{

        Configuration configuration = new Configuration();
        configuration.setDefaultEncoding("utf-8");
        // 模板存放路径
        configuration.setClassForTemplateLoading(this.getClass(), "/templates/code");
        // 获取模板文件
        Template template = configuration.getTemplate("UserInfo.ftl");
        // 模拟数据
        List<Study> studyList =new ArrayList<>();
        studyList.add(new Study("gaolei","root","1","user","123456"));
        studyList.add(new Study("zhangsan","123","15","admin","12232"));

        Map<String,Object> resultMap = new HashMap<>();
        resultMap.put("studyList",studyList);
        File outFile = new File("UserInfoTest.doc");
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),"UTF-8"));
        try {
            template.process(resultMap,out);
            out.flush();
            out.close();
            return null;
        } catch (TemplateException e) {
            e.printStackTrace();
        }
        return ResultBean.success();
    }

    @RequestMapping("user/requireInfo")
    @ResponseBody
    public ResultBean  userRequireInfo() throws  IOException{
        Configuration configuration = new Configuration();
        configuration.setDefaultEncoding("utf-8");
        configuration.setClassForTemplateLoading(this.getClass(), "/templates/code");
        Template template = configuration.getTemplate("need.ftl");
        Map<String , Object> resultMap = new HashMap<>();
        List<UserInfo> userInfoList = new ArrayList<>();
        userInfoList.add(new UserInfo("2019","安全环保处质量安全科2608室","风险研判","9:30","10:30","风险研判","风险研判原型设计","参照甘肃分公司提交的分析研判表，各个二级单位维护自己的风险研判信息，需要一个简单的风险上报流程，各个二级单位可以看到所有的分析研判信息作为一个知识成果共享。","张三","李四"));
        resultMap.put("userInfoList",userInfoList);
        File outFile = new File("userRequireInfo.doc");
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),"UTF-8"));
        try {
            template.process(resultMap,out);
            out.flush();
            out.close();
            return null;
        } catch (TemplateException e) {
            e.printStackTrace();
        }
        return ResultBean.success();
    }
    /**
     * 导出带图片的Word
     * @throws IOException
     * @return
     */
    @RequestMapping("user/exportPic")
    @ResponseBody
    public ResultBean exportPic() throws IOException {
        Configuration configuration = new Configuration();
        configuration.setDefaultEncoding("utf-8");
        configuration.setClassForTemplateLoading(this.getClass(), "/templates/code");
        Template template = configuration.getTemplate("userPic.ftl");
        Map<String,Object> map = new HashMap<>();
        map.put("name","gaolei");
        map.put("date","2015-10-12");
        map.put("imgCode",imageToString());
        File outFile = new File("userWithPicture.doc");
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),"UTF-8"));
        try {
            template.process(map,out);
            out.flush();
            out.close();
            return null;
        } catch (TemplateException e) {
            e.printStackTrace();
        }
        return  ResultBean.success();
    }

    /**
     * 图片流转base64字符串
     * */
    public String imageToString() {
        URL resource = BaseController.class.getClassLoader().getResource("static/img/a.png");
        String imgFile = resource.getPath();
        InputStream in;
        byte[] data = null;
        try {
            in = new FileInputStream(imgFile);
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        String imageCodeBase64 =  Base64Utils.encodeToString(data);

        return imageCodeBase64;
    }
}
