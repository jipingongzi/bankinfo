package com.xiaozhen.bankinfo;

import com.xiaozhen.bankinfo.common.ExcelUtil;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.Map;

/**
 * @author xz
 */
@RestController
@RequestMapping("/rest")
public class RestService {

    @GetMapping("/login")
    public Map<String,String> login(@RequestParam("role")String role){
        return null;

    }

    @GetMapping("/data/export")
    public void siteChecklist(HttpServletResponse response) throws IOException {
        ExcelUtil.excelResponse(ExcelUtil.data(),response,"数据统计");
    }
}
