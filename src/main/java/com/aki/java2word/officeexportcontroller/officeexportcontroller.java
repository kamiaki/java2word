package com.aki.java2word.officeexportcontroller;

import com.kmood.datahandle.DocumentProducer;
import com.kmood.utils.FileUtils;
import lombok.extern.slf4j.Slf4j;
import main.Main;
import org.apache.commons.lang3.StringUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.net.URL;
import java.util.*;

@Controller
@Slf4j
public class officeexportcontroller {

    /**
     * 文本输出
     *
     * @param response
     * @throws Exception
     */
    @RequestMapping(value = "test1")
    public void testTextOutModel(HttpServletResponse response) throws Exception {
        ClassLoader classLoader = ClassLoader.getSystemClassLoader();
        // ActualModelPath 和 xmlPath 只要有一个有值就能找到模板，最好都写
        String ActualModelPath = classLoader.getResource("./model/").toURI().getPath();
        String xmlPath = classLoader.getResource("./model").toURI().getPath();
        // 组装参数
        Map<String, Object> map = new HashMap<>();
        map.put("text", "kmood-文本占位输出");
        map.put("text1", "kmood-文本占位输出2");
        //  自定义配置
        //  DocumentProducer dp = new DocumentProducer(FreemarkerUtil.configuration, ActualModelPath);
        DocumentProducer dp = new DocumentProducer(ActualModelPath);
        String modelFile = dp.Complie(xmlPath, "text.xml", false);
        log.info("使用的模板文件:{}", modelFile);
        // 设置强制下载不打开 设置文件名
        response.setContentType("application/force-download");
        response.addHeader("Content-Disposition", StringUtils.join("attachment;fileName=", new Date(), ".doc"));
        dp.produce(map, response.getOutputStream());
    }


    /**
     * 循环输出测试
     *
     * @param response
     * @throws Exception
     */
    @RequestMapping(value = "test2")
    public void testTextLoopOutModel(HttpServletResponse response) throws Exception {
        ClassLoader classLoader = ClassLoader.getSystemClassLoader();
        String ActualModelPath = classLoader.getResource("./model/").toURI().getPath();
        String xmlPath = classLoader.getResource("./model").toURI().getPath();
        // 组装参数 不能用this那种方式创建
        Map<String, Object> map = new HashMap<>();
        map.put("zzdhm", "kmood-制造单号码");
        map.put("ydwcrq", "kmood-预定完成日期");
        map.put("cpmc", "kmood-产品名称");
        map.put("jyrq", "kmood-交运日期");
        map.put("sl", "kmood-数量");
        map.put("xs", "kmood-箱数");
        List<Object> zxsmList = new ArrayList<>();
        Map<String, Object> zxsmmap = new HashMap<>();
        zxsmmap.put("xh", "kmood-箱号");
        zxsmmap.put("xs", "kmood-箱数");
        zxsmmap.put("zrl", "kmood-梅香");
        zxsmmap.put("zsl", "kmood-交运日期");
        zxsmmap.put("sm", "kmood-交运日期");
        zxsmList.add(zxsmmap);
        Map<String, Object> zxsmmap1 = new HashMap<>();
        zxsmmap1.put("xh", "kmood-制造单号码");
        zxsmmap1.put("xs", "kmood-预定完成日期");
        zxsmmap1.put("zrl", "kmood-产品名称");
        zxsmmap1.put("zsl", "kmood-交运日期");
        zxsmmap1.put("sm", "kmood-交运日期");
        zxsmList.add(zxsmmap1);
        map.put("zxsm", zxsmList);
        map.put("sbsm", "kmood-商标说明");
        map.put("bt", "kmood OfficeExport 导出word");
        DocumentProducer dp = new DocumentProducer(ActualModelPath);
        String modelFile = dp.Complie(xmlPath, "包装说明表（范例A）.xml", true);
        log.info("使用的模板文件:{}", modelFile);
        // 设置强制下载不打开 设置文件名
        response.setContentType("application/force-download");
        response.addHeader("Content-Disposition", StringUtils.join("attachment;fileName=", new Date(), ".doc"));
        dp.produce(map, response.getOutputStream());
    }

    /**
     * 输出带图片的
     *
     * @param response
     * @throws Exception
     */
    @RequestMapping(value = "test3")
    public void testImageOutModel(HttpServletResponse response) throws Exception {
        ClassLoader classLoader = ClassLoader.getSystemClassLoader();
        String ActualModelPath = classLoader.getResource("./model/").toURI().getPath();
        String xmlPath = classLoader.getResource("./model").toURI().getPath();
        HashMap<String, Object> map = new HashMap<>();
        //读取输出图片
        URL introUrl = classLoader.getResource("./picture/exportTestPicture-intro.png");
        URL codeUrl = classLoader.getResource("./picture/exportTestPicture-code.png");
        URL titleUrl = classLoader.getResource("./picture/exportTestPicture-title.png");
        map.put("intro", Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(introUrl.toURI().getPath())));
        map.put("code", Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(codeUrl.toURI().getPath())));
        map.put("title", Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(titleUrl.toURI().getPath())));
        //编译输出
        DocumentProducer dp = new DocumentProducer(ActualModelPath);
        String modelFile = dp.Complie(xmlPath, "picture.xml", true);
        log.info("使用的模板文件:{}", modelFile);
        // 设置强制下载不打开 设置文件名
        response.setContentType("application/force-download");
        response.addHeader("Content-Disposition", StringUtils.join("attachment;fileName=", new Date(), ".doc"));
        dp.produce(map, response.getOutputStream());
    }
}
