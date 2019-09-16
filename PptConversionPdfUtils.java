package cn.spacecg.web.utils;

import lombok.extern.log4j.Log4j;
import org.apache.commons.io.IOUtils;
import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.util.Random;

/**
 * @ClassName: PptConversionPdfUtils
 * @Description: TODO 类的描述
 * @author: 符纯涛
 * @date: 2019/8/24
 */
@Log4j
public class PptConversionPdfUtils {

    /**
     * @param args
     */
    private static OfficeManager officeManager;
    //openoffice4安装路径
    private static String SERVER_PATH = "/opt/openoffice4/";
//    private static String SERVER_PATH = "C:\\Program Files (x86)\\OpenOffice 4";
    private static int port[] = {8100};
    File toFile;

    /**
     * pdf 转换器
     *
     * @param inputFile ppt文件
     * @throws IOException
     */
    public String pdfConverter(MultipartFile inputFile, HttpServletRequest request) throws Exception {

        try {
            // 开启服务器
            startService();
            String folderName = FileUtil.typeName(4);
//            String realPath = "/opt/html/"+folderName;
            String realPath = request.getSession().getServletContext().getRealPath(folderName);
            File targetFile = new File(realPath);
            if (!targetFile.exists()) {
                targetFile.mkdirs();
            }
            // 获取Excel文件
            final File pptFile = getFile(inputFile);
            String suffix = ".pdf";
            String name = String.valueOf(System.currentTimeMillis());
//            final File pdfFile = File.createTempFile(getRandByNum(6), suffix,targetFile);
            final File pdfFile = File.createTempFile(name, suffix,targetFile);
            // Office文档转换器
            OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
            converter.convert(pptFile, pdfFile);
            log.info("openoffice4转换pdf成功......");
            // 返回
            toFile = pdfFile;
            // 删除
            deleteFile(pptFile);
//            deleteFile(pdfFile);
            String pdfPath = pdfFile.getPath();
            int pdf = pdfPath.indexOf("pdf");
            String substring = pdfPath.substring(pdf);
            substring = substring.replaceAll("\\\\", "/");
            log.info("openoffice4转换pdf的路径：{}"+ substring);
            return substring;
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
            log.info("准备启动openoffice4服务....");
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
            log.info("office转换服务启动成功!");
        } catch (Exception ce) {
            log.error("office转换服务启动失败!详细信息:" + ce);
        }
    }

    /**
     * 关闭服务器
     */
    private static void stopService() {
        log.info("关闭office转换服务....");
        if (officeManager != null) {
            officeManager.stop();
        }
        log.info("关闭office转换成功!");
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
