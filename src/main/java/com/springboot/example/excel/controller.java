package com.springboot.example.excel;

import com.springboot.example.excel.entity.User;
import lombok.extern.slf4j.Slf4j;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;

@RestController
@RequestMapping(value = "/")
@Slf4j
public class controller {

    @RequestMapping(value = "/export", method = RequestMethod.GET)
    public void excel(HttpServletResponse response) {

        /**
         * excel导出
         * 1.获取数据集List 插入到map集合中
         * 2.根据模板生成新的excel
         * 3.将新生成的excel文件从浏览器输出
         * 4.删除新生成的模板文件
         */
        List<User> list = new ArrayList();
        list.add(new User(1, "zs", 21, new Date()));
        list.add(new User(2, "ls", 22, new Date()));
        Map<String, Object> beans = new HashMap();
        beans.put("list", list);

        //加载excel模板文件
        File file = null;
        try {
            file = ResourceUtils.getFile("classpath:excel/aaa.xlsx");
        } catch (FileNotFoundException e) {
            log.error("template file path cannot be found");
        }

        //配置下载路径
        String path = "/download/";
        createDir(new File(path));

        //根据模板生成新的excel
        File excelFile = createNewFile(beans, file, path);

        //浏览器端下载文件
        downloadFile(response, excelFile);

        //删除服务器生成文件
        deleteFile(excelFile);

    }


    /**
     * 根据excel模板生成新的excel
     *
     * @param beans
     * @param file
     * @param path
     * @return
     */
    private File createNewFile(Map<String, Object> beans, File file, String path) {
        XLSTransformer transformer = new XLSTransformer();

        //可以写工具类来生成命名规则
        String name = "bbb.xlsx";
        File newFile = new File(path + name);


        try (InputStream in = new BufferedInputStream(new FileInputStream(file));
             OutputStream out = new FileOutputStream(newFile)) {
            Workbook workbook = transformer.transformXLS(in, beans);
            workbook.write(out);
            out.flush();
            return newFile;
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
        return newFile;
    }

    /**
     * 将服务器新生成的excel从浏览器下载
     *
     * @param response
     * @param excelFile
     */
    private void downloadFile(HttpServletResponse response, File excelFile) {
        /* 设置文件ContentType类型，这样设置，会自动判断下载文件类型 */
        response.setContentType("multipart/form-data");
        /* 设置文件头：最后一个参数是设置下载文件名 */
        response.setHeader("Content-Disposition", "attachment;filename=" + excelFile.getName());
        try (
                InputStream ins = new FileInputStream(excelFile);
                OutputStream os = response.getOutputStream()
        ) {
            byte[] b = new byte[1024];
            int len;
            while ((len = ins.read(b)) > 0) {
                os.write(b, 0, len);
            }
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
    }

    /**
     * 浏览器下载完成之后删除服务器生成的文件
     * 也可以设置定时任务去删除服务器文件
     *
     * @param excelFile
     */
    private void deleteFile(File excelFile) {
        excelFile.delete();
    }

    //如果目录不存在创建目录 存在则不创建
    private void createDir(File file) {
        if (!file.exists()) {
            file.mkdirs();
        }
    }
}
