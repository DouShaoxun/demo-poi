package cn.dsx.demopoi;

import cn.dsx.demopoi.utils.DrawImageUtils;
import cn.dsx.demopoi.utils.ExcelUtils;
import cn.dsx.demopoi.utils.SnowflakeIdWorker;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigDecimal;
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

    /**
     * https://www.cnblogs.com/acm-bingzi/p/poiPicture.html
     * @throws Exception
     */
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
        //XSSFWorkbook workbook = new XSSFWorkbook(in);
        Workbook workbook = WorkbookFactory.create(in);


        //读取了模板内所有sheet内容
        Sheet sheet = workbook.getSheetAt(0);
        // sheet只能获取一个
        Drawing patriarch = sheet.createDrawingPatriarch();

        /**
         * https://www.cnblogs.com/dtts/p/4741575.html
         * EXCEL列高度的单位是磅,Apache POI的行高度单位是缇(twip)
         * 　　DPI = 1英寸内可显示的像素点个数。通常电脑屏幕是96DPI, IPhone4s的屏幕是326DPI, 普通激光黑白打印机是400DPI
         * 　　要计算POI行高或者Excel的行高，就先把它行转换到英寸，再乘小DPI就可以得到像素
         * 　　像素= (Excel的行高度/72)*DPI
         * 所以获取行高的像素值的方法就是： (row.getHeightInPoints() / 72) * 96
         * 像素 ＝ (磅/72)*DPI
         *
         * 像素= (Excel的行高度/72)*DPI
         *
         * 像素= (POI中的行高/20/72)*DPI
         *
         * Excel的行高度=像素/DPI*72
         *
         * POI中的行高=像素/DPI*72*20
         */
        //Cell cell = sheet.getRow(2).getCell(2);//获取第4行 第三列
        //int columnIndex = cell.getColumnIndex();
        //int rowIndex = cell.getRowIndex();
        //System.out.println(ExcelUtils.isMergedRegion(sheet, rowIndex, columnIndex));
        ////float columnWidthInPixels = sheet.getColumnWidthInPixels(columnIndex); //  单位不是像素，是1/256个字符宽度  3.8版本没有此方法
        //int columnWidth = sheet.getColumnWidth(columnIndex);//单位是像素
        //
        //float heightInPoints = cell.getRow().getHeightInPoints();//  获取的是excel行高的榜值
        //float heightInPointsPoi = cell.getRow().getHeightInPoints() / 72 * 96;//poi高度单位计算

        //======================  ======================//

        //System.out.println(ExcelUtils.isMergedRegion(sheet, 2, 2));
        //System.out.println(ExcelUtils.isMergedRegion(sheet, 0, 0));
        //CellRangeAddress mergedRegion = ExcelUtils.getMergedRegion(sheet, 2, 2);
        //CellRangeAddress mergedRegion1 = ExcelUtils.getMergedRegion(sheet, 3, 3);
        //CellRangeAddress mergedRegion2 = ExcelUtils.getMergedRegion(sheet, 0, 2);
        //int numberOfCells = mergedRegion1.getNumberOfCells();// 获取合并单元格当中 单元格数量
        //System.out.println(numberOfCells);
        ////
        //// 循环计算 合并单元格 高度和宽度
        //int totalHeight = 0;
        //for (int row = mergedRegion.getFirstRow(); row <= mergedRegion.getLastRow(); row++) {
        //    totalHeight += sheet.getRow(row).getHeightInPoints();
        //}
        //System.out.println("totalHeight:" + totalHeight);
        //
        //double totalWeight = 0;
        //double totalWeightMillimetres;
        //for (int col = mergedRegion.getFirstColumn(); col <= mergedRegion.getLastColumn(); col++) {
        //    totalWeight += sheet.getColumnWidth(col);
        //}
        //totalWeightMillimetres = ExcelUtils.ConvertImageUnits.widthUnits2Millimetres((short) totalWeight);

        //====================== 0.jpg ======================//
        ByteArrayOutputStream byteArrayOut_0 = new ByteArrayOutputStream();
        log.info(imagePath + "/0.jpg");
        File image_0 = new File(imagePath + "/0.jpg");
        BufferedImage user_headImg_0 = DrawImageUtils.drawImage(image_0);
        ImageIO.write(user_headImg_0, "jpg", byteArrayOut_0);
        int height_0 = user_headImg_0.getHeight();// 图片高度
        int widt_0 = user_headImg_0.getWidth();// 图片宽度
        BigDecimal imageRatioCanvas = ratioCanvas(widt_0, height_0);// 图片比例
        // 获取合并单元格
        CellRangeAddress mergedRegion_0 = ExcelUtils.getMergedRegion(sheet, 2, 2);



        // 循环计算 合并单元格 高度和宽度
        int totalHeight_0 = 0;
        for (int row = mergedRegion_0.getFirstRow(); row <= mergedRegion_0.getLastRow(); row++) {
            totalHeight_0 += sheet.getRow(row).getHeightInPoints();
        }
        double cellHeight = totalHeight_0 / 72 * 96;//像素
        double totalHeightMillimetres = ExcelUtils.ConvertImageUnits.pointsToMillimeters(totalHeight_0);
        double totalWidth = 0;
        double cellWidth = 0;
        double totalWeightMillimetres;
        for (int col = mergedRegion_0.getFirstColumn(); col <= mergedRegion_0.getLastColumn(); col++) {
            totalWidth += sheet.getColumnWidth(col);
            cellWidth += sheet.getColumnWidthInPixels(col);//
        }
        totalWeightMillimetres = ExcelUtils.ConvertImageUnits.widthUnits2Millimetres((short) totalWidth);
        BigDecimal cellRatioCanvas = ratioCanvas(totalWeightMillimetres, totalHeightMillimetres);// 单元格比例

        boolean flagType = false; //缩放类型
        //长宽比例的系数
        double a = 1.0;
        double b = 1.0;
        double standardWidth = 112;
        double standardHeight = 41;
        double needWeightMillimetres = 0D;
        double needHeightMillimetres = 0D;
        XSSFClientAnchor anchor_0 = new XSSFClientAnchor();
        if (imageRatioCanvas.compareTo(cellRatioCanvas) >= 0) {
            // 图片过宽 根据图片的宽和单元格的宽比进行缩放
            System.out.println("图片过宽 根据图片的宽和单元格的宽比进行缩放");
            // 计算缩放比例
            //if (cellWidth > widt_0) {
            //    //a = widt_0 / cellWidth;
            //    a = standardWidth / cellWidth;
            //} else {
            //    //a = cellWidth / widt_0;
            //    a = cellWidth / standardWidth;
            //}
            a = standardWidth / cellWidth;
            flagType = true;
            needHeightMillimetres = Math.abs(totalWeightMillimetres / imageRatioCanvas.doubleValue());
            int needRowNum = 0;
            double hasHeightMM = 0D;
            for (int row = mergedRegion_0.getFirstRow(); row <= mergedRegion_0.getLastRow(); row++) {
                // 寻找适合的行坐标
                if (hasHeightMM >= needHeightMillimetres) {
                    break;
                }
                hasHeightMM += ExcelUtils.ConvertImageUnits.pointsToMillimeters((short) sheet.getRow(row).getHeightInPoints());
                needRowNum++;
            }
            double spaceHeightMM = hasHeightMM - needHeightMillimetres;// 留白部分
            double rowCoordinatesPerMM = 0.0D;
            // rowHeightMM 真实毫米高度
            double rowHeightMM = ExcelUtils.ConvertImageUnits.pointsToMillimeters((short) sheet.getRow(mergedRegion_0.getFirstRow() + needRowNum - 1).getHeightInPoints());

            // 每毫米多少像素
            rowCoordinatesPerMM = ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS / rowHeightMM;
            int pictureHeightCoordinates = 0;
            // 留白像素值
            pictureHeightCoordinates = (int) (spaceHeightMM * rowCoordinatesPerMM);

            // 计算偏移位置
            int i = (mergedRegion_0.getLastRow() - mergedRegion_0.getFirstRow() - needRowNum) / 2;//左右留白
            int dx1 = 100 * Units.EMU_PER_PIXEL;
            int dy1 = 50 * Units.EMU_PER_PIXEL;
            int dx2 = (ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS - 100) * Units.EMU_PER_PIXEL;
            int dy2 = (ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - pictureHeightCoordinates - 50) * Units.EMU_PER_PIXEL;
            int col1 = mergedRegion_0.getFirstColumn();
            int row1 = mergedRegion_0.getFirstRow();
            int col2 = mergedRegion_0.getLastColumn();
            int row2 = mergedRegion_0.getLastRow() + needRowNum - 1 + i;
            //anchor_0.setDx1(dx1);
            //anchor_0.setDy1(dy1);
            //anchor_0.setDx2(dx2);
            //anchor_0.setDy2(dy2);
            anchor_0.setCol1(col1);
            anchor_0.setRow1(row1);
            anchor_0.setCol2(col2);
            anchor_0.setRow2(row2);

        } else {
            //
            System.out.println("图片过高 根据图片的高和单元格的高比进行缩放");

            if (cellHeight > height_0) {
                b = height_0 / cellHeight;
            } else {
                b = cellHeight / height_0;
            }


            needWeightMillimetres = Math.abs(totalHeightMillimetres * imageRatioCanvas.doubleValue());
            int needColNum = 0;
            double hasWeightMM = 0D;

            for (int col = mergedRegion_0.getFirstColumn(); col <= mergedRegion_0.getLastColumn(); col++) {
                if (hasWeightMM >= needWeightMillimetres) {
                    break;
                }
                hasWeightMM += ExcelUtils.ConvertImageUnits.widthUnits2Millimetres(
                        (short) sheet.getColumnWidth(col));
                needColNum++;
            }

            double spaceWeightMM = hasWeightMM - needWeightMillimetres;
            double colCoordinatesPerMM = 0.0D;
            double colWidthMM = ExcelUtils.ConvertImageUnits.widthUnits2Millimetres((short) sheet.getColumnWidth(mergedRegion_0.getFirstColumn() + needColNum - 1));

            colCoordinatesPerMM = ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS / colWidthMM;
            int pictureWidthCoordinates = 0;
            pictureWidthCoordinates = (int) (spaceWeightMM * colCoordinatesPerMM);

            int i = 0;
            if (needColNum <= mergedRegion_0.getLastColumn() - mergedRegion_0.getFirstColumn() + 1) {
                i = (mergedRegion_0.getLastColumn() - mergedRegion_0.getFirstColumn() - needColNum + 1) / 2;
            }

            int dx1 = 100 * Units.EMU_PER_PIXEL;
            int dy1 = 50 * Units.EMU_PER_PIXEL;
            int dx2 = (ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS - pictureWidthCoordinates - 100) * Units.EMU_PER_PIXEL;
            int dy2 = (ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - 50) * Units.EMU_PER_PIXEL;
            int col1 = mergedRegion_0.getFirstColumn() + i;
            int row1 = (mergedRegion_0.getFirstRow());
            int col2 = (mergedRegion_0.getFirstColumn() + needColNum - 1 + i);
            int row2 = (mergedRegion_0.getLastRow());

            //anchor_0.setDx1(dx1);
            //anchor_0.setDy1(dy1);
            //anchor_0.setDx2(dx2);
            //anchor_0.setDy2(dy2);
            anchor_0.setCol1(col1);
            anchor_0.setRow1(row1);
            anchor_0.setCol2(col2);
            anchor_0.setRow2(row2);

        }
        Picture picture_0 = patriarch.createPicture(anchor_0, workbook.addPicture(byteArrayOut_0.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        if (flagType) {
            picture_0.resize(a);
        } else {
            picture_0.resize(b);
        }



        // 设置图片的属性
        //int col1_0 = 2;
        //int row1_0 = 2;
        //int col2_0 = 22;
        //int row2_0 = 13;
        //XSSFClientAnchor anchor_0 = new XSSFClientAnchor(0, 0, 0, 0 * Units.EMU_PER_PIXEL, col1_0, row1_0, col2_0, row2_0);


        //  Sets the anchor type
        //anchor_0.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);//3.15
        //anchor_0.setAnchorType(DONT_MOVE_AND_RESIZE);//3.8
        // 插入图片 
        //Picture picture_0 = patriarch.createPicture(anchor_0, workbook.addPicture(byteArrayOut_0.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));\
        //picture_0.resize(1, 1);// 设置缩放比例
        //====================== 0.jpg ======================//
/*

        //====================== 1.jpg ======================//
        //ExcelTransferUtils.getMergedRegionmergedRegion_0(sheet, titleElement.getRowIndex(), titleElement.getColIndex());
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
        ClientAnchor anchor_1 = new XSSFClientAnchor(0, 0, 150 * Units.EMU_PER_PIXEL, 150000 * Units.EMU_PER_PIXEL, col1_1, row1_1, col2_1, row2_1);
        //  Sets the anchor type
        anchor_1.setAnchorType(DONT_MOVE_AND_RESIZE);
        // 插入图片 
        patriarch.createPicture(anchor_1, workbook.addPicture(byteArrayOut_1.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        //====================== 1.jpg ======================//
*/
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
        XSSFClientAnchor regionr = (XSSFClientAnchor) patriarch.createAnchor(0, 0, 150 * Units.EMU_PER_PIXEL, 150, 0, 0, 50, 50);
        regionr.setAnchorType(3);
        XSSFSimpleShape region1Shapevr = ((XSSFDrawing) patriarch).createSimpleShape(regionr);
        region1Shapevr.setShapeType(ShapeTypes.LINE);


        // 操作xlsx ClassCastException: org.apache.poi.xssf.usermodel.XSSFClientAnchor cannot be cast to org.apache.poi.hssf.usermodel.HSSFClientAnchor
        //HSSFClientAnchor regionr = (HSSFClientAnchor) patriarch.createAnchor(0, 0, 150 * Units.EMU_PER_PIXEL, 150, 0, 0, 50, 50);
        //regionr.setAnchorType(3);
        //HSSFSimpleShape region1Shapevr = ((HSSFPatriarch) patriarch).createSimpleShape(regionr);
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


    /**
     * 计算宽高比例
     * 宽：高
     *
     * @return
     */
    public static BigDecimal ratioCanvas(double w, double h) {
        BigDecimal wigth = new BigDecimal(w);
        BigDecimal height = new BigDecimal(h);
        return wigth.divide(height, 5, BigDecimal.ROUND_HALF_UP);
    }

}
