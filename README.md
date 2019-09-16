# Office转换成pdf

## 1、依赖

- guava

```xml
<dependency>
            <groupId>com.google.guava</groupId>
            <artifactId>guava</artifactId>xmlxx
            <version>19.0</version>
</dependency>
```

- livesense

```xml
<dependency>
            <groupId>com.github.livesense</groupId>
            <artifactId>jodconverter-core</artifactId>
            <version>1.0.5</version>
</dependency>
```

## 2、本机测试

- 本机安装一个openoffice软件即可。

　　若是被部署项目的服务器，可以在服务器本地安装一个openoffice软件；也可以在其他服调用其他服务器上的openoffice服务进行文档转换。

​		openoffice的下载地址：http://www.openoffice.org/

## 3、中文乱码

- Linux参考

https://blog.csdn.net/fanjin287659245/article/details/80360767

Windows不会乱码，乱码一般字体缺失。

### 4、代码

- DemoController.java

```java 
package com.fude.demo.controller

import com.fude.demo.util.PptConversionPdfUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;

@Controller
public class DemoController {

    @ResponseBody
    @RequestMapping(value = "/getPDF", method = RequestMethod.POST)
    public File getPDF(MultipartFile inputFile) throws Exception {
        PptConversionPdfUtils pptConversionPdfUtils = new PptConversionPdfUtils();
        return pptConversionPdfUtils.pdfConverter(inputFile);
    }
}
```



- PptConversionPdfUtils.java

```java
package com.fude.demo.util;

import org.apache.commons.io.IOUtils;
import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.Random;

/**
 * ppt转换pdf
 */
public class PptConversionPdfUtils {
    /**
     * @param args
     */
    private static OfficeManager officeManager;
    private static String SERVER_PATH = "/opt/openoffice4/";
    private static int port[] = {8100};
    File toFile;

    /**
     * pdf 转换器
     *
     * @param inputFile ppt文件
     * @throws IOException
     */
    public File pdfConverter(MultipartFile inputFile) throws Exception {

        try {
            // 开启服务器
            startService();
            // 获取Excel文件
            final File pptFile = getFile(inputFile);
            final File pdfFile = File.createTempFile(getRandByNum(6), ".pdf");
            // Office文档转换器
            OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
            converter.convert(pptFile, pdfFile);
            // 返回
            toFile = pdfFile;
            // 删除
            deleteFile(pptFile);
//            deleteFile(pdfFile);
            return toFile;
        } finally {
            // 关闭服务器
            stopService();
        }

    }


    /**
     * 打开服务器
     */
    private static void startService() {
        DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
        try {
            System.out.println("准备启动服务....");
            // 设置OpenOffice.org安装目录
            configuration.setOfficeHome(SERVER_PATH);
            // 设置转换端口，默认为8100
            configuration.setPortNumbers(port);
            // 设置任务执行超时为5分钟
            configuration.setTaskExecutionTimeout(1000 * 60 * 5L);
            // 设置任务队列超时为24小时
            configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);

            officeManager = configuration.buildOfficeManager();
            // 启动服务
            officeManager.start();
            System.out.println("office转换服务启动成功!");
        } catch (Exception ce) {
            System.out.println("office转换服务启动失败!详细信息:" + ce);
        }
    }

    /**
     * 关闭服务器
     */
    private static void stopService() {
        System.out.println("关闭office转换服务....");
        if (officeManager != null) {
            officeManager.stop();
        }
        System.out.println("关闭office转换成功!");
    }

    /**
     * 获取File文件流对象
     *
     * @param inputFile
     * @return
     * @throws IOException
     */
    private File getFile(MultipartFile inputFile) throws IOException {
        File toFile = null;
        if (inputFile.equals("") || inputFile.getSize() <= 0) {
            inputFile = null;
        } else {
            InputStream ins = null;
            ins = inputFile.getInputStream();
            toFile = new File(inputFile.getOriginalFilename());
            inputStreamToFile(ins, toFile);
            ins.close();
        }
        return toFile;
    }

    /**
     * 产生num位的随机数
     *
     * @return
     */
    private static synchronized String getRandByNum(int num) {
        // 定制长度
        String length = "1";
        for (int i = 0; i < num; i++) {
            length += "0";
        }
        // 获取随机数
        Random rad = new Random();
        String result = rad.nextInt(Integer.parseInt(length)) + "";

        if (result.length() != num) {
            return getRandByNum(num);
        }
        return result;
    }

    /**
     * File 转 MultipartFile
     * @param file
     * @throws Exception
     */
    public static MultipartFile fileToMultipartFile(File file) throws Exception {

        FileInputStream fileInput = new FileInputStream(file);
        MultipartFile toMultipartFile = new MockMultipartFile("file",file.getName(),"text/plain", IOUtils.toByteArray(fileInput));
        toMultipartFile.getInputStream();
        return toMultipartFile;
    }

    /**
     * 转换器
     *
     * @param ins
     * @param file
     */
    private static void inputStreamToFile(InputStream ins, File file) {
        try {
            OutputStream os = new FileOutputStream(file);
            int bytesRead = 0;
            byte[] buffer = new byte[8192];
            while ((bytesRead = ins.read(buffer, 0, 8192)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
            os.close();
            ins.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 删除
     *
     * @param files
     */
    private void deleteFile(File... files) {
        for (File file : files) {
            if (file.exists()) {
                file.delete();
            }
        }
    }
}

```

