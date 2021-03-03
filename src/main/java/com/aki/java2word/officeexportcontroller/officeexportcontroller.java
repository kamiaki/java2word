package com.aki.java2word.officeexportcontroller;

import com.kmood.datahandle.DocumentProducer;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

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
        String ActualModelPath = classLoader.getResource("./model/").getPath();
        String xmlPath = classLoader.getResource("./model").getPath();
        // 组装参数
        HashMap<String, Object> map = new HashMap<>();
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
     * 文本循环输出
     *
     * @param response
     * @throws Exception
     */
    @RequestMapping(value = "test2")
    public void testTextLoopOutModel(HttpServletResponse response) throws Exception {
        ClassLoader classLoader = ClassLoader.getSystemClassLoader();
        String ActualModelPath = classLoader.getResource("./model/").getPath();
        String xmlPath = classLoader.getResource("./model").getPath();
        // 组装参数
        HashMap<String, Object> map = new HashMap<>();
        map.put("zzdhm", "kmood-制造单号码");
        map.put("ydwcrq", "kmood-预定完成日期");
        map.put("cpmc", "kmood-产品名称");
        map.put("jyrq", "kmood-交运日期");
        map.put("sl", "kmood-数量");
        map.put("xs", "kmood-箱数");

        ArrayList<Object> zxsmList = new ArrayList<>();
        HashMap<String, Object> zxsmmap = new HashMap<>();
        zxsmmap.put("xh", "kmood-箱号");
        zxsmmap.put("xs", "kmood-箱数");
        zxsmmap.put("zrl", "kmood-梅香");
        zxsmmap.put("zsl", "kmood-交运日期");
        zxsmmap.put("sm", "kmood-交运日期");
        zxsmList.add(zxsmmap);
        HashMap<String, Object> zxsmmap1 = new HashMap<>();
        zxsmmap1.put("xh", "kmood-制造单号码");
        zxsmmap1.put("xs", "kmood-预定完成日期");
        zxsmmap1.put("zrl","kmood-产品名称");
        zxsmmap1.put("zsl","kmood-交运日期");
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
}
