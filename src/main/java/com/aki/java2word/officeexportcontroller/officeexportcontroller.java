package com.aki.java2word.officeexportcontroller;

import com.kmood.datahandle.DocumentProducer;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.util.Date;
import java.util.HashMap;

@Controller
@Slf4j
public class officeexportcontroller {

    /**
     * 文本输出
     * @param response
     * @throws Exception
     */
    @RequestMapping(value = "test1")
    public void testTextOutModel(HttpServletResponse response) throws Exception {
        ClassLoader classLoader = ClassLoader.getSystemClassLoader();
        // ActualModelPath 和 xmlPath 只要有一个有值就能找到模板，最好都写
        String ActualModelPath = classLoader.getResource("./model/").getPath();
        String xmlPath = classLoader.getResource("./model").getPath();
        HashMap<String, Object> map = new HashMap<>();
        map.put("text", "kmood-文本占位输出");
        map.put("text1", "kmood-文本占位输出2");
        DocumentProducer dp = new DocumentProducer(ActualModelPath);
        String modelFile = dp.Complie(xmlPath, "text.xml", false);
        log.info("使用的模板文件:{}", modelFile);
        // 设置强制下载不打开 设置文件名
        response.setContentType("application/force-download");
        response.addHeader("Content-Disposition", StringUtils.join("attachment;fileName=",new Date(),".doc"));
        dp.produce(map, response.getOutputStream());
    }
}
