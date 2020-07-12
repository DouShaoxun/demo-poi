package cn.dsx.demopoi;

import cn.dsx.demopoi.utils.DrawImageUtils;
import cn.dsx.demopoi.utils.SnowflakeIdWorker;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

import static org.apache.poi.ss.usermodel.ClientAnchor.DONT_MOVE_AND_RESIZE;

/**
 * 文档
 * Sets the y coordinate within the second cell Note - XSSF and HSSF have a slightly different coordinate system,
 * values in XSSF are larger by a factor of Units.EMU_PER_PIXEL
 * https://poi.apache.org/apidocs/4.0/org/apache/poi/xssf/usermodel/XSSFClientAnchor.html
 *
 * @Classname: PoiTest
 * @Author: Dsx
 * @Date: 2020/07/10/22:57
 */
@SpringBootTest
@Slf4j
public class XLSXPoiTest {

    @Autowired
    SnowflakeIdWorker idGener;
    private static SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMddHHmmss");

    @Test
    void contextLoads() throws Exception {
        // 模板路径
        File directory = new File("src/main/resources");
        String courseFile = directory.getCanonicalPath();
        String excelPath = courseFile + "/templates/excel";
        String filePath = excelPath + "/template.xlsx";
        log.info("xlsx模板路径：" + filePath);
        String imagePath = courseFile + "/static/image";
        log.info("image根路径：" + imagePath);

        // 读入excel文件
        File file = new File(filePath);
        InputStream in = new FileInputStream(file);


        //读取excel模板
        XSSFWorkbook workbook = new XSSFWorkbook(in);


        //读取了模板内所有sheet内容
        XSSFSheet sheet = workbook.getSheetAt(0);
        // sheet只能获取一个
        XSSFDrawing patriarch = sheet.createDrawingPatriarch();





        //====================== 0.jpg ======================//
        ByteArrayOutputStream byteArrayOut_0 = new ByteArrayOutputStream();
        log.info(imagePath + "/0.jpg");
        File image_0 = new File(imagePath + "/0.jpg");
        BufferedImage user_headImg_0 = DrawImageUtils.drawImage(image_0);
        ImageIO.write(user_headImg_0, "jpg", byteArrayOut_0);
        // 设置图片的属性
        int col1_0 = 1;
        int row1_0 = 11;
        int col2_0 = 5;
        int row2_0 = 20;
        XSSFClientAnchor anchor_0 = new XSSFClientAnchor(0, 0, 0, 0, col1_0, row1_0, col2_0, row2_0);
        //  Sets the anchor type
        anchor_0.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        // 插入图片 
        patriarch.createPicture(anchor_0, workbook.addPicture(byteArrayOut_0.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        //====================== 0.jpg ======================//




        //====================== 1.jpg ======================//
        //ExcelTransferUtils.getMergedRegionPositionRange(sheet, titleElement.getRowIndex(), titleElement.getColIndex());
        ByteArrayOutputStream byteArrayOut_1 = new ByteArrayOutputStream();
        log.info(imagePath + "/1.jpg");
        File image_1 = new File(imagePath + "/1.jpg");
        BufferedImage user_headImg_1 = DrawImageUtils.drawImage(image_1);
        ImageIO.write(user_headImg_1, "jpg", byteArrayOut_1);
        // 设置图片的属性
        int col1_1 = 25;
        int row1_1 = 2;
        int col2_1 = 25;
        int row2_1 = 13;
        ClientAnchor anchor_1 = new XSSFClientAnchor(0, 0, 150 * Units.EMU_PER_PIXEL, 150000* Units.EMU_PER_PIXEL, col1_1, row1_1, col2_1, row2_1);
        //  Sets the anchor type
        anchor_1.setAnchorType(DONT_MOVE_AND_RESIZE);
        // 插入图片 
        patriarch.createPicture(anchor_1, workbook.addPicture(byteArrayOut_1.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        //====================== 1.jpg ======================//

        /*


        //====================== 2.jpg ======================//
        ByteArrayOutputStream byteArrayOut_2 = new ByteArrayOutputStream();
        log.info(imagePath + "/2.jpg");
        File image_2 = new File(imagePath + "/2.jpg");
        BufferedImage user_headImg_2 =  DrawImageUtils.drawImage(image_2);
        ImageIO.write(user_headImg_2, "jpg", byteArrayOut_2);
        // 设置图片的属性
        int col1_2 = 2;
        int row1_2 = 17;
        int col2_2 = 22;
        int row2_2 = 22;
        XSSFClientAnchor anchor_2 = new XSSFClientAnchor(0, 0, 0, 0, col1_2, row1_2, col2_2, row2_2);
        //  Sets the anchor type
        anchor_2.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        // 插入图片 
        patriarch.createPicture(anchor_2, workbook.addPicture(byteArrayOut_2.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        //====================== 2.jpg ======================//



        //====================== 3.jpg ======================//
        ByteArrayOutputStream byteArrayOut_3 = new ByteArrayOutputStream();
        log.info(imagePath + "/3.jpg");
        File image_3 = new File(imagePath + "/3.jpg");
        BufferedImage user_headImg_3 = DrawImageUtils.drawImage(image_3);
        ImageIO.write(user_headImg_3, "jpg", byteArrayOut_3);
        // 设置图片的属性
        int col1_3 = 10;
        int row1_3 = 11;
        int col2_3 = 15;
        int row2_3 = 20;
        XSSFClientAnchor anchor_3 = new XSSFClientAnchor(0, 0, 0, 0, col1_3, row1_3, col2_3, row2_3);
        //  Sets the anchor type
        anchor_3.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        // 插入图片 
        patriarch.createPicture(anchor_3, workbook.addPicture(byteArrayOut_3.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        //====================== 3.jpg ======================//


        //====================== 4.jpg ======================//
        ByteArrayOutputStream byteArrayOut_4 = new ByteArrayOutputStream();
        log.info(imagePath + "/4.jpg");
        File image_4 = new File(imagePath + "/4.jpg");
        BufferedImage user_headImg_4 = DrawImageUtils.drawImage(image_4);
        ImageIO.write(user_headImg_4, "jpg", byteArrayOut_4);
        // 设置图片的属性
        int col1_4 = 10;
        int row1_4 = 11;
        int col2_4 = 15;
        int row2_4 = 20;
        XSSFClientAnchor anchor_4 = new XSSFClientAnchor(0, 0, 0, 0, col1_4, row1_4, col2_4, row2_4);
        //  Sets the anchor type
        anchor_4.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        // 插入图片 
        patriarch.createPicture(anchor_4, workbook.addPicture(byteArrayOut_4.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        //====================== 4.jpg ======================//


        //====================== large.jpg ======================//
        ByteArrayOutputStream byteArrayOut_large = new ByteArrayOutputStream();
        log.info(imagePath + "/large.jpg");
        File image_large = new File(imagePath + "/large.jpg");
        BufferedImage user_headImg_large = DrawImageUtils.drawImage(image_large);
        ImageIO.write(user_headImg_large, "jpg", byteArrayOut_large);
        // 设置图片的属性
        int col1_large = 10;
        int row1_large = 11;
        int col2_large = 15;
        int row2_large = 20;
        XSSFClientAnchor anchor_large = new XSSFClientAnchor(0, 0, 0, 0, col1_large, row1_large, col2_large, row2_large);
        //  Sets the anchor type
        anchor_large.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        // 插入图片 
        patriarch.createPicture(anchor_large, workbook.addPicture(byteArrayOut_large.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        //====================== large.jpg ======================//


        //====================== middle.jpg ======================//
        ByteArrayOutputStream byteArrayOut_middle = new ByteArrayOutputStream();
        log.info(imagePath + "/middle.jpg");
        File image_middle = new File(imagePath + "/middle.jpg");
        BufferedImage user_headImg_middle = DrawImageUtils.drawImage(image_middle);
        ImageIO.write(user_headImg_middle, "jpg", byteArrayOut_middle);
        // 设置图片的属性
        int col1_middle = 10;
        int row1_middle = 11;
        int col2_middle= 15;
        int row2_middle = 20;
        XSSFClientAnchor anchor_middle = new XSSFClientAnchor(0, 0, 0, 0, col1_middle, row1_middle, col2_middle, row2_middle);
        //  Sets the anchor type
        anchor_middle.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        // 插入图片 
        patriarch.createPicture(anchor_middle, workbook.addPicture(byteArrayOut_middle.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        //====================== middle.jpg ======================//


        //====================== small.jpg ======================//
        ByteArrayOutputStream byteArrayOut_small = new ByteArrayOutputStream();
        log.info(imagePath + "/small.jpg");
        File image_small = new File(imagePath + "/small.jpg");
        BufferedImage user_headImg_small = DrawImageUtils.drawImage(image_small);
        ImageIO.write(user_headImg_small, "jpg", byteArrayOut_small);
        // 设置图片的属性
        int col1_small = 10;
        int row1_small = 11;
        int col2_small= 15;
        int row2_small = 20;
        XSSFClientAnchor anchor_small = new XSSFClientAnchor(0, 0, 0, 0, col1_small, row1_small, col2_small, row2_small);
        //  Sets the anchor type
        anchor_small.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        // 插入图片 
        patriarch.createPicture(anchor_small, workbook.addPicture(byteArrayOut_small.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        //====================== small.jpg ======================//



        */


        // 画线 此处3.15无效   3.8版本可以 原因待查
        // https://blog.csdn.net/Czhou9468/article/details/103789940
        XSSFClientAnchor regionr = patriarch.createAnchor(0, 0, 150 * Units.EMU_PER_PIXEL, 150 , 0, 0, 0, 5);
        regionr.setAnchorType(3);
        XSSFSimpleShape region1Shapevr = patriarch.createSimpleShape(regionr);
        region1Shapevr.setShapeType(ShapeTypes.LINE);


        //HSSFPatriarch drawingPatriarch = sheet.createDrawingPatriarch();
        //HSSFClientAnchor  regionr = drawingPatriarch.createAnchor(9525*10, 9525*10, 150 * Units.EMU_PER_PIXEL, 150*9525, 0, 0,  50, 50);
        //HSSFSimpleShape region1Shapevr = patriarch.createSimpleShape(regionr);
        //region1Shapevr.setShapeType(ShapeTypes.LINE);

        workbook.setActiveSheet(0);
        //修改模板内容导出新模板
        String exprotPath = excelPath + "/exprot/";
        File dir = new File(exprotPath);
        if (!dir.exists()) {
            dir.mkdir();
        }
        String format = simpleDateFormat.format(new Date());
        String outputFilePath = exprotPath + format + ".xlsx";
        log.info("输出路径：" + outputFilePath);


        FileOutputStream outputStream = new FileOutputStream(outputFilePath);
        workbook.write(outputStream);

        outputStream.close();
    }

}
