package cn.dsx.demopoi;

import cn.dsx.demopoi.utils.DrawImageUtils;
import cn.dsx.demopoi.utils.ExcelUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

@SpringBootTest
@Slf4j
public class PoiTest {
    private static SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMddHHmmss");

    /**
     * https://www.cnblogs.com/acm-bingzi/p/poiPicture.html
     *
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

        // 坐标
        int[][] coordinate = {
                {2, 2},         // 区域1
                {17, 2},        // 区域2
                {25, 8},        // 区域3
                {2, 25},        // 区域4
                {18, 25},       // 区域5
                {42, 2},        // 区域6
                {44, 2},        // 区域7
                {3, 27},        // 区域8
                {42, 27},       // 区域9
                {42, 25}        // 区域10
        };

        for (int i = 0; i < coordinate.length; i++) {
            //buildExcelImage(imagePath + "/0.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage(imagePath + "/1.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage(imagePath + "/2.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage(imagePath + "/3.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage(imagePath + "/4.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage(imagePath + "/large.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage(imagePath + "/middle.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage(imagePath + "/small.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
        }
        for (int i = 0; i < coordinate.length; i++) {
            //buildExcelImage2(imagePath + "/0.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage2(imagePath + "/1.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage2(imagePath + "/2.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage2(imagePath + "/3.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage2(imagePath + "/4.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage2(imagePath + "/large.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage2(imagePath + "/middle.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage2(imagePath + "/small.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
        }
        for (int i = 0; i < coordinate.length; i++) {
            //buildExcelImage3(imagePath + "/0.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage3(imagePath + "/1.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage3(imagePath + "/2.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage3(imagePath + "/3.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage3(imagePath + "/4.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage3(imagePath + "/large.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage3(imagePath + "/middle.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage3(imagePath + "/small.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
        }
        for (int i = 0; i < coordinate.length; i++) {
            //buildExcelImage4(imagePath + "/0.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage4(imagePath + "/1.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage4(imagePath + "/2.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage4(imagePath + "/3.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage4(imagePath + "/4.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage4(imagePath + "/large.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage4(imagePath + "/middle.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);
            //buildExcelImage4(imagePath + "/small.jpg", sheet, patriarch, workbook, coordinate[i][0], coordinate[i][1]);

        }
        // 单独测试
        int k = 8;
        //buildExcelImage3(imagePath + "/1.jpg", sheet, patriarch, workbook, coordinate[k][0], coordinate[k][1]);
        buildExcelImage4(imagePath + "/1.jpg", sheet, patriarch, workbook, coordinate[k][0], coordinate[k][1]);

        // 画线 此处3.15无效   3.8版本可以 原因待查
        // https://blog.csdn.net/Czhou9468/article/details/103789940
        XSSFClientAnchor regionr = (XSSFClientAnchor) patriarch.createAnchor(0, 0, 150 * Units.EMU_PER_PIXEL, 150, 0, 0, 50, 50);
        regionr.setAnchorType(3);
        XSSFSimpleShape region1Shapevr = ((XSSFDrawing) patriarch).createSimpleShape(regionr);
        region1Shapevr.setShapeType(ShapeTypes.LINE);
        region1Shapevr.setFillColor(255, 0, 0);
        region1Shapevr.setLineWidth(12000);

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
     * 参考
     * https://www.cnblogs.com/acm-bingzi/p/poiPicture.html
     * https://blog.csdn.net/xlxin/article/details/72726032?utm_medium=distribute.pc_relevant_t0.none-task-blog-BlogCommendFromMachineLearnPai2-1.compare&depth_1-utm_source=distribute.pc_relevant_t0.none-task-blog-BlogCommendFromMachineLearnPai2-1.compare
     *
     * @param imagePath 图片路径
     * @param sheet
     * @param patriarch
     * @param workbook
     * @param firstRow  单元格所在行
     * @param firstCol  单元格所在列
     * @throws IOException
     */
    public static void buildExcelImage(String imagePath, Sheet sheet, Drawing patriarch, Workbook workbook, int firstRow, int firstCol) throws IOException {
        ByteArrayOutputStream byteArrayOut_0 = new ByteArrayOutputStream();
        log.info(imagePath);
        System.out.println(firstRow + "," + firstCol);
        File image_0 = new File(imagePath);
        BufferedImage user_headImg_0 = DrawImageUtils.drawImage(image_0);
        ImageIO.write(user_headImg_0, "jpg", byteArrayOut_0);
        int imgY = user_headImg_0.getHeight();// 图片高度
        int imgX = user_headImg_0.getWidth();// 图片宽度
        // 获取合并单元格
        CellRangeAddress mergedRegion_0 = ExcelUtils.getMergedRegion(sheet, firstRow, firstCol);
        // 循环计算 合并单元格 高度和宽度
        int totalHeight_0 = 0;
        for (int row = mergedRegion_0.getFirstRow(); row <= mergedRegion_0.getLastRow(); row++) {
            totalHeight_0 += sheet.getRow(row).getHeightInPoints();
        }
        double cellWidth = 0;
        for (int col = mergedRegion_0.getFirstColumn(); col <= mergedRegion_0.getLastColumn(); col++) {
            cellWidth += sheet.getColumnWidthInPixels(col);
        }
        // 计算偏移量
        double cellX = Units.pixelToPoints((int) cellWidth);
        double cellY = totalHeight_0;
        int[] anchorArray = {0, 0, 0, 0};
        anchorArray = ExcelUtils.calCellAnchor(cellX, cellY, imgX, imgY);


        XSSFClientAnchor anchor_0 = new XSSFClientAnchor();
        int col1 = mergedRegion_0.getFirstColumn();
        int row1 = mergedRegion_0.getFirstRow();
        int col2 = mergedRegion_0.getLastColumn();
        int row2 = mergedRegion_0.getLastRow();
        anchor_0.setDx1(anchorArray[0]);
        anchor_0.setDy1(anchorArray[1]);
        anchor_0.setDx2(anchorArray[2]);
        anchor_0.setDy2(anchorArray[3]);
        anchor_0.setCol1(col1);
        anchor_0.setRow1(row1);
        anchor_0.setCol2(col2 + 1);// 设置结束单元格为右下角相邻的单元格 然后 对应的偏移量设置为负数
        anchor_0.setRow2(row2 + 1);
        anchor_0.setAnchorType(ClientAnchor.MOVE_AND_RESIZE);


        Picture picture_0 = patriarch.createPicture(anchor_0, workbook.addPicture(byteArrayOut_0.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        //picture_0.resize(1, 0.5);
        //picture_0.resize(0.5, 1);//resize是按照单元格长度来计算图片长度  所以此方法不会等比缩放

    }


    public static void buildExcelImage2(String imagePath, Sheet sheet, Drawing patriarch, Workbook workbook, int firstRow, int firstCol) throws IOException {
        ByteArrayOutputStream byteArrayOut_0 = new ByteArrayOutputStream();
        XSSFClientAnchor anchor_0 = new XSSFClientAnchor();
        anchor_0.setAnchorType(XSSFClientAnchor.MOVE_AND_RESIZE);//移动
        log.info(imagePath);
        System.out.println(firstRow + "," + firstCol);
        File image_0 = new File(imagePath);
        BufferedImage user_headImg_0 = DrawImageUtils.drawImage(image_0);
        ImageIO.write(user_headImg_0, "jpg", byteArrayOut_0);
        int imgY = user_headImg_0.getHeight();// 图片高度
        int imgX = user_headImg_0.getWidth();// 图片宽度
        // 获取合并单元格
        CellRangeAddress mergedRegion_0 = ExcelUtils.getMergedRegion(sheet, firstRow, firstCol);

        // 循环计算 合并单元格 高度和宽度         begin
        int totalHeight_0 = 0;
        for (int row = mergedRegion_0.getFirstRow(); row <= mergedRegion_0.getLastRow(); row++) {
            //System.out.println(sheet.getRow(row).getHeightInPoints());
            //返回以点大小度量的行高。如果未设置高度，则返回默认工作表值
            // 16.5/72*96= 22像素
            totalHeight_0 += sheet.getRow(row).getHeightInPoints();
        }
        ;
        // 转换成毫米
        //double totalHeightMillimetres = ExcelUtils.ConvertImageUnits.pointsToMillimeters(totalHeight_0);
        double totalHeightMillimetres = ExcelUtils.ConvertImageUnits.pointsToMillimeters(totalHeight_0);


        double cellWidth = 0;
        double totalWidth = 0;
        for (int col = mergedRegion_0.getFirstColumn(); col <= mergedRegion_0.getLastColumn(); col++) {
            // 12.4444275  获取像素
            cellWidth += sheet.getColumnWidthInPixels(col);
            totalWidth += sheet.getColumnWidth(col); //单位不是像素，是1/256个字符宽度
        }
        //象素数bai / DPI = 英寸数
        //英寸数 * 25.4 = 毫米du数
        // 转换成毫米
        double totalWeightMillimetres = cellWidth / 96 * 25.4;
        // 循环计算 合并单元格 高度和宽度         end

        BigDecimal cellRatioCanvas = ratioCanvas(totalWeightMillimetres, totalHeightMillimetres);   // 单元格比例
        BigDecimal imageRatioCanvas = ratioCanvas(imgX, imgY);                                      // 图片比例
        double needWeightMillimetres = 0D;
        double needHeightMillimetres = 0D;


        if (imageRatioCanvas.compareTo(cellRatioCanvas) >= 0) {
            // x:y ：w：h
            //x*h=y*w
            //h=y*w/x =w*y/x

            // 图片过宽 根据图片的宽和单元格的宽比进行缩放
            System.out.println("根据宽度缩放");
            int needRowNum = 0;
            double hasHeightMM = 0D;
            needHeightMillimetres = Math.abs(totalWeightMillimetres / imageRatioCanvas.doubleValue());
            for (int row = mergedRegion_0.getFirstRow(); row <= mergedRegion_0.getLastRow(); row++) {
                if (hasHeightMM >= needHeightMillimetres) {
                    break;
                }
                hasHeightMM += ExcelUtils.ConvertImageUnits.pointsToMillimeters((short) sheet.getRow(row).getHeightInPoints());
                needRowNum++;
            }

            double spaceHeightMM = hasHeightMM - needHeightMillimetres;//计算空白
            double rowCoordinatesPerMM = 0.0D;
            Row row = sheet.getRow(mergedRegion_0.getFirstRow() + needRowNum - 1);
            float heightInPoints = row.getHeightInPoints();
            double rowHeightMM = ExcelUtils.ConvertImageUnits.pointsToMillimeters((short) heightInPoints);

            //每个单元格宽度分成1023份  高分成256份
            rowCoordinatesPerMM = ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS / rowHeightMM;//每毫米站多少个单位
            int pictureHeightCoordinates = 0;
            pictureHeightCoordinates = (int) (spaceHeightMM * rowCoordinatesPerMM);// 需要留白的毫米数量乘上单位 得到偏移量

            // 计算偏移位置
            int i = (mergedRegion_0.getLastRow() - mergedRegion_0.getFirstRow() - needRowNum) / 2;//左右留白
            int dx1 = 5;
            int dy1 = 5;
            int dx2 = (ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS - 5);
            int dy2 = (ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - pictureHeightCoordinates);
            int col1 = mergedRegion_0.getFirstColumn();
            int row1 = mergedRegion_0.getFirstRow();
            int col2 = mergedRegion_0.getLastColumn();
            int row2 = mergedRegion_0.getFirstRow() + needRowNum - 1 + i;

            anchor_0.setDx1((int) Math.round(dx1 * XSSFShape.EMU_PER_PIXEL));
            anchor_0.setDy1((int) Math.round(dy1 * XSSFShape.EMU_PER_PIXEL));
            anchor_0.setDx2((int) Math.round(dx2 * XSSFShape.EMU_PER_PIXEL));
            anchor_0.setDy2((int) Math.round(dy2 * XSSFShape.EMU_PER_PIXEL));

            anchor_0.setCol1(col1);
            anchor_0.setRow1(row1);
            anchor_0.setCol2(col2);
            anchor_0.setRow2(row2);
            System.out.println("dx1 ：" + anchor_0.getDx1());
            System.out.println("dy1 ：" + anchor_0.getDy1());
            System.out.println("dx2 ：" + anchor_0.getDx2());
            System.out.println("dy2 ：" + anchor_0.getDy2());
            System.out.println("col1：" + anchor_0.getCol1());
            System.out.println("row1：" + anchor_0.getRow1());
            System.out.println("col2：" + anchor_0.getCol2());
            System.out.println("row2：" + anchor_0.getRow2());

        } else {

            System.out.println("根据高度缩放");
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
            int dx1 = 10;
            int dy1 = 10;
            int dx2 = (ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS - pictureWidthCoordinates - 10);
            int dy2 = (ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - 10);
            int col1 = mergedRegion_0.getFirstColumn() + i;
            int row1 = (mergedRegion_0.getFirstRow());
            int col2 = (mergedRegion_0.getFirstColumn() + needColNum - 1 + i);
            int row2 = (mergedRegion_0.getLastRow());

            //anchor_0.setDx1((int) Math.round(dx1 * XSSFShape.EMU_PER_PIXEL));
            //anchor_0.setDy1((int) Math.round(dy1 * Units.EMU_PER_POINT));
            //anchor_0.setDx2((int) Math.round(dx2 * XSSFShape.EMU_PER_PIXEL));
            //anchor_0.setDy2((int) Math.round(dy2 * Units.EMU_PER_POINT));

            anchor_0.setCol1(col1);
            anchor_0.setRow1(row1);
            anchor_0.setCol2(col2);
            anchor_0.setRow2(row2);

            System.out.println("dx1 ：" + anchor_0.getDx1());
            System.out.println("dy1 ：" + anchor_0.getDy1());
            System.out.println("dx2 ：" + anchor_0.getDx2());
            System.out.println("dy2 ：" + anchor_0.getDy2());
            System.out.println("col1：" + anchor_0.getCol1());
            System.out.println("row1：" + anchor_0.getRow1());
            System.out.println("col2：" + anchor_0.getCol2());
            System.out.println("row2：" + anchor_0.getRow2());

        }

        Picture picture_0 = patriarch.createPicture(anchor_0, workbook.addPicture(byteArrayOut_0.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));

    }


    /**
     * todo
     *
     * @param imagePath
     * @param sheet
     * @param patriarch
     * @param workbook
     * @param firstRow
     * @param firstCol
     * @throws IOException
     */
    public static void buildExcelImage3(String imagePath, Sheet sheet, Drawing patriarch, Workbook workbook, int firstRow, int firstCol) throws IOException {
        ByteArrayOutputStream byteArrayOut_0 = new ByteArrayOutputStream();
        XSSFClientAnchor anchor_0 = new XSSFClientAnchor();
        anchor_0.setAnchorType(XSSFClientAnchor.MOVE_AND_RESIZE);//移动
        log.info("图片路径：" + imagePath);
        log.info("(" + firstRow + "," + firstCol + ")");
        File image_0 = new File(imagePath);
        BufferedImage user_headImg_0 = DrawImageUtils.drawImage(image_0);
        ImageIO.write(user_headImg_0, "jpg", byteArrayOut_0);


        //======================================获取图片属性  begin====================================

        int imgY = user_headImg_0.getHeight();// 图片高度
        int imgX = user_headImg_0.getWidth();// 图片宽度

        // 图片尺寸转换成毫米
        double imgXPoint = Units.pixelToPoints((int) imgX);
        double imgXMM = ExcelUtils.ConvertImageUnits.pointsToMillimeters(imgXPoint);
        double imgYPoint = Units.pixelToPoints((int) imgY);
        double imgYMM = ExcelUtils.ConvertImageUnits.pointsToMillimeters(imgYPoint);

        //======================================获取图片属性  end====================================


        // 获取合并单元格
        CellRangeAddress mergedRegion_0 = ExcelUtils.getMergedRegion(sheet, firstRow, firstCol);

        //======================================循环计算 合并单元格 高度和宽度  begin====================================


        int totalHeight_0 = 0;
        for (int row = mergedRegion_0.getFirstRow(); row <= mergedRegion_0.getLastRow(); row++) {
            //返回以点大小度量的行高。如果未设置高度，则返回默认工作表值
            // 16.5/72*96= 22像素
            totalHeight_0 += sheet.getRow(row).getHeightInPoints();
        }
        // 表格高度 转换成毫米
        double cellYMM = ExcelUtils.ConvertImageUnits.pointsToMillimeters(totalHeight_0);


        double cellWidth = 0;
        for (int col = mergedRegion_0.getFirstColumn(); col <= mergedRegion_0.getLastColumn(); col++) {
            // 12.4444275  获取像素
            cellWidth += sheet.getColumnWidthInPixels(col);
        }
        // 像素数 / DPI = 英寸数
        // 英寸数 * 25.4 = 毫米数
        // 表格宽度 转换成毫米
        double cellXMM = ExcelUtils.ConvertImageUnits.pointsToMillimeters(Units.pixelToPoints((int) cellWidth));

        //======================================循环计算 合并单元格 高度和宽度  end====================================


        double needXMillimetres = 0D;
        double needYMillimetres = 0D;

        boolean createFlag = false;
        // todo
        // 图片和单元格总共 又四种情况  均采用毫米单位进行比较
        // 1. imgXMM > cellXMM   &&  imgYMM  <= cellYMM    图片宽度大于单元格宽度 且图片高度小于单元格高度
        // 2. imgXMM <= cellXMM   &&  imgYMM  > cellYMM    图片高度大于单元格高度 且图片宽度小于单元格宽度
        // 3. imgXMM > cellXMM   &&  imgYMM  > cellYMM    图片宽高均大于单元格
        // 4. imgXMM < cellXMM   &&  imgYMM  < cellYMM    图片宽高均小于单元格
        if (imgXMM <= cellXMM && imgYMM > cellYMM) { // 过高，需要缩放高度  对应情况2
            log.info("imgXMM <= cellXMM && imgYMM > cellYMM");
            // TODO: 2020/8/27

            //// 1.先判断图片的宽度和单元格的宽度之比
            //double widthRatio = imgXMM / cellXMM;
            //
            //// 图片过宽 根据图片的宽和单元格的宽比进行缩放
            //int needRowNum = 0;
            //double hasHeightMM = 0D;
            ////needHeightMillimetres = Math.abs(totalWeightMillimetres / imageRatioCanvas.doubleValue());
            //needHeightMillimetres = Math.abs(imgYMM / widthRatio);
            //for (int row = mergedRegion_0.getFirstRow(); row <= mergedRegion_0.getLastRow(); row++) {
            //    if (hasHeightMM >= needHeightMillimetres) {
            //        break;
            //    }
            //    hasHeightMM += ExcelUtils.ConvertImageUnits.pointsToMillimeters((short) sheet.getRow(row).getHeightInPoints());
            //    needRowNum++;
            //}
            //
            //double spaceHeightMM = hasHeightMM - needHeightMillimetres;//计算空白
            //double rowCoordinatesPerMM = 0.0D;
            //Row row = sheet.getRow(mergedRegion_0.getFirstRow() + needRowNum - 1);
            //float heightInPoints = row.getHeightInPoints();
            //double rowHeightMM = ExcelUtils.ConvertImageUnits.pointsToMillimeters((short) heightInPoints);
            //
            ////每个单元格宽度分成1023份  高分成256份
            //rowCoordinatesPerMM = ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS / rowHeightMM;//每毫米站多少个单位
            //int pictureHeightCoordinates = 0;
            //pictureHeightCoordinates = (int) (spaceHeightMM * rowCoordinatesPerMM);// 需要留白的毫米数量乘上单位 得到偏移量
            //
            //// 计算偏移位置
            //int i = (mergedRegion_0.getLastRow() - mergedRegion_0.getFirstRow() - needRowNum) / 2;//左右留白
            //int dx1 = 0;
            //int dy1 = 0;
            //int dx2 = (ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS - 0);
            //int dy2 = (ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - pictureHeightCoordinates - 0);
            //int col1 = mergedRegion_0.getFirstColumn();
            //int row1 = mergedRegion_0.getFirstRow();
            //int col2 = mergedRegion_0.getLastColumn();
            //int row2 = mergedRegion_0.getFirstRow() + needRowNum - 1 + i;
            //
            //
            //anchor_0.setDx1((int) Math.round(dx1 * Units.EMU_PER_PIXEL));
            //anchor_0.setDy1((int) Math.round(dy1 * Units.EMU_PER_PIXEL));
            //anchor_0.setDx2((int) Math.round(dx2 * Units.EMU_PER_PIXEL));
            //anchor_0.setDy2((int) Math.round(dy2 * Units.EMU_PER_PIXEL));
            //
            //anchor_0.setCol1(col1);
            //anchor_0.setRow1(row1);
            //anchor_0.setCol2(col2);
            //anchor_0.setRow2(row2);

        } else if (imgXMM > cellXMM && imgYMM <= cellYMM) { // 过宽 缩放宽度  对应情况1
            // TODO: 2020/8/27
            log.info("imgXMM > cellXMM && imgYMM <= cellYMM");
            //double heightRatio = imgYMM / cellYMM;
            //double ratio = 1.0;
            //needWeightMillimetres = Math.abs(imgXMM / heightRatio);
            //int needColNum = 0;
            //double hasWeightMM = 0D;
            //for (int col = mergedRegion_0.getFirstColumn(); col <= mergedRegion_0.getLastColumn(); col++) {
            //    if (hasWeightMM >= needWeightMillimetres) {
            //        break;
            //    }
            //    hasWeightMM += ExcelUtils.ConvertImageUnits.widthUnits2Millimetres((short) sheet.getColumnWidth(col));
            //    needColNum++;
            //}
            //
            //double spaceWeightMM = hasWeightMM - needWeightMillimetres;
            //double colCoordinatesPerMM = 0.0D;
            //double colWidthMM = ExcelUtils.ConvertImageUnits.widthUnits2Millimetres((short) sheet.getColumnWidth(mergedRegion_0.getFirstColumn() + needColNum - 1));
            //
            //colCoordinatesPerMM = ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS / colWidthMM;
            //int pictureWidthCoordinates = 0;
            //pictureWidthCoordinates = (int) (spaceWeightMM * colCoordinatesPerMM);
            //
            //
            //int i = 0;
            //if (needColNum <= mergedRegion_0.getLastColumn() - mergedRegion_0.getFirstColumn() + 1) {
            //    i = (mergedRegion_0.getLastColumn() - mergedRegion_0.getFirstColumn() - needColNum + 1) / 2;
            //}
            //int dx1 = 0;
            //int dy1 = 0;
            //int dx2 = (ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS - pictureWidthCoordinates - 0);
            //int dy2 = (ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - 0);
            //
            //
            //int col1 = mergedRegion_0.getFirstColumn() + i;
            //int row1 = (mergedRegion_0.getFirstRow());
            //int col2 = (mergedRegion_0.getFirstColumn() + needColNum - 1 + i);
            //int row2 = (mergedRegion_0.getLastRow());
            //
            //anchor_0.setDx1((int) Math.round(dx1 * Units.EMU_PER_PIXEL));
            //anchor_0.setDy1((int) Math.round(dy1 * Units.EMU_PER_PIXEL));
            //anchor_0.setDx2((int) Math.round(dx2 * Units.EMU_PER_PIXEL));
            //anchor_0.setDy2((int) Math.round(dy2 * Units.EMU_PER_PIXEL));
            //
            //anchor_0.setCol1(col1);
            //anchor_0.setRow1(row1);
            //anchor_0.setCol2(col2);
            //anchor_0.setRow2(row2);

        } else if (imgXMM < cellXMM && imgYMM < cellYMM) {
            // TODO: 2020/8/27
            log.info("imgXMM < cellXMM && imgYMM < cellYMM");
            //
            //needWeightMillimetres = imgXMM;// 计算需要的宽高 单位毫米  暂时采用原尺寸
            //needHeightMillimetres = imgYMM;
            //
            ////  =====================计算宽所在的列
            //int needColNum = 0;
            //double hasWeightMM = 0D;
            //for (int col = mergedRegion_0.getFirstColumn(); col <= mergedRegion_0.getLastColumn(); col++) {
            //    if (hasWeightMM >= needWeightMillimetres) {
            //        break;
            //    }
            //    hasWeightMM += ExcelUtils.ConvertImageUnits.widthUnits2Millimetres((short) sheet.getColumnWidth(col));
            //    needColNum++;
            //}
            //
            //// 计算横轴空白
            //double spaceWeightMM = hasWeightMM - needWeightMillimetres;
            //double colCoordinatesPerMM = 0.0D;
            //// 获得宽所在的列
            //double colWidthMM = ExcelUtils.ConvertImageUnits.widthUnits2Millimetres((short) sheet
            //        .getColumnWidth(mergedRegion_0.getFirstColumn() + needColNum - 1));
            //
            //colCoordinatesPerMM = ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS / colWidthMM;
            //int pictureWidthCoordinates = (int) (spaceWeightMM * colCoordinatesPerMM);
            ////  =====================
            //
            //
            //// 计算纵轴空白
            //int needRowNum = 0;
            //double hasHeightMM = 0D;
            //for (int row = mergedRegion_0.getFirstRow(); row <= mergedRegion_0.getLastRow(); row++) {
            //    if (hasHeightMM >= needHeightMillimetres) {
            //        break;
            //    }
            //    hasHeightMM += ExcelUtils.ConvertImageUnits.pointsToMillimeters((short) sheet.getRow(row).getHeightInPoints());
            //    needRowNum++;
            //}
            //
            //double spaceHeightMM = hasHeightMM - needHeightMillimetres;//计算空白  单位毫米
            //double rowCoordinatesPerMM = 0.0D;
            //Row row = sheet.getRow(mergedRegion_0.getFirstRow() + needRowNum - 1);
            //float heightInPoints = row.getHeightInPoints();
            //double rowHeightMM = ExcelUtils.ConvertImageUnits.pointsToMillimeters((short) heightInPoints);
            //
            ////每个单元格宽度分成1023份  高分成256份
            //rowCoordinatesPerMM = ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS / rowHeightMM;//每毫米站多少个单位
            //int pictureHeightCoordinates = (int) (spaceHeightMM * rowCoordinatesPerMM);// 需要留白的毫米数量乘上单位 得到偏移量
            //
            //int dx1 = 0;
            //int dy1 = 0;
            //int dx2 = (ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS - pictureWidthCoordinates - 0);
            //int dy2 = (ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - pictureHeightCoordinates - 0);
            //
            //
            //int col1 = mergedRegion_0.getFirstColumn();
            //int row1 = (mergedRegion_0.getFirstRow());
            //int col2 = (mergedRegion_0.getFirstColumn() + needColNum - 1);
            //int row2 = mergedRegion_0.getFirstRow() + needRowNum - 1;
            //
            //anchor_0.setDx1((int) Math.round(dx1 * Units.EMU_PER_PIXEL));
            //anchor_0.setDy1((int) Math.round(dy1 * Units.EMU_PER_PIXEL));
            //anchor_0.setDx2((int) Math.round(dx2 * Units.EMU_PER_PIXEL));
            //anchor_0.setDy2((int) Math.round(dy2 * Units.EMU_PER_PIXEL));
            //
            //anchor_0.setCol1(col1);
            //anchor_0.setRow1(row1);
            //anchor_0.setCol2(col2);
            //anchor_0.setRow2(row2);


        } else if (imgXMM >= cellXMM && imgYMM >= cellYMM) {

            log.info("imgXMM >= cellXMM && imgYMM >= cellYMM");
            double widthRatio = imgXMM / cellXMM;//宽
            double heightRatio = imgYMM / cellYMM;//高
            double ratio = widthRatio > heightRatio ? widthRatio : heightRatio;  // 根据大的来缩放


            needXMillimetres = Math.abs(imgXMM / ratio);// 计算需要的宽高 单位毫米
            needYMillimetres = Math.abs(imgYMM / ratio);

            //  =====================计算宽所在的列
            int needColNum = 0;
            double hasXMM = 0D;
            for (int col = mergedRegion_0.getFirstColumn(); col <= mergedRegion_0.getLastColumn(); col++) {
                if (hasXMM >= needXMillimetres) {
                    break;
                }
                float columnWidthInPixels = sheet.getColumnWidthInPixels(col);
                hasXMM += ExcelUtils.ConvertImageUnits.pointsToMillimeters(Units.pixelToPoints((int) columnWidthInPixels));
                needColNum++;
            }

            // 计算横轴空白
            double spaceWeightMM = hasXMM - needXMillimetres;
            double colCoordinatesPerMM = 0.0D;
            // 获得宽所在的列
            int columnWidthPixel = (int) sheet.getColumnWidthInPixels(mergedRegion_0.getFirstColumn() + needColNum - 1);
            double colWidthMM = ExcelUtils.ConvertImageUnits.pointsToMillimeters(Units.pixelToPoints((int) columnWidthPixel));


            colCoordinatesPerMM = ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS / colWidthMM;
            int pictureWidthCoordinates = (int) (spaceWeightMM * colCoordinatesPerMM);
            //  =====================


            // 计算纵轴空白
            int needRowNum = 0;
            double hasHeightMM = 0D;
            for (int row = mergedRegion_0.getFirstRow(); row <= mergedRegion_0.getLastRow(); row++) {
                if (hasHeightMM >= needYMillimetres) {
                    break;
                }
                hasHeightMM += ExcelUtils.ConvertImageUnits.pointsToMillimeters((short) sheet.getRow(row).getHeightInPoints());
                needRowNum++;
            }

            double spaceHeightMM = hasHeightMM - needYMillimetres;//计算空白  单位毫米
            double rowCoordinatesPerMM = 0.0D;
            Row row = sheet.getRow(mergedRegion_0.getFirstRow() + needRowNum - 1);
            double rowHeightMM = ExcelUtils.ConvertImageUnits.pointsToMillimeters((short) row.getHeightInPoints());

            //每个单元格宽度分成1023份  高分成256份
            rowCoordinatesPerMM = ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS / rowHeightMM;//每毫米站多少个单位
            int pictureHeightCoordinates = (int) (spaceHeightMM * rowCoordinatesPerMM);// 需要留白的毫米数量乘上单位 得到偏移量


            int spaceMM = 0;// 四周留白
            int dx1 = spaceMM;
            int dy1 = spaceMM;
            int dx2 = ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS - pictureWidthCoordinates - spaceMM;
            int dy2 = ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - pictureHeightCoordinates - spaceMM;

            //int dx2 = (ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS - 0);
            //int dy2 = (ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - 26 - 0);


            int col1 = mergedRegion_0.getFirstColumn();
            int row1 = mergedRegion_0.getFirstRow();
            int col2 = mergedRegion_0.getFirstColumn() + needColNum - 1;
            int row2 = mergedRegion_0.getFirstRow() + needRowNum - 1;

            anchor_0.setDx1(Math.round(dx1 * XSSFShape.EMU_PER_PIXEL));
            anchor_0.setDy1(Math.round(dy1 * XSSFShape.EMU_PER_PIXEL));
            anchor_0.setDx2(Math.round(dx2 * XSSFShape.EMU_PER_PIXEL));
            anchor_0.setDy2(Math.round(dy2 * XSSFShape.EMU_PER_PIXEL));

            anchor_0.setCol1(col1);
            anchor_0.setRow1(row1);
            anchor_0.setCol2(col2);
            anchor_0.setRow2(row2);
            createFlag = true;
        }

        log.info("dx1 ：" + anchor_0.getDx1());
        log.info("dy1 ：" + anchor_0.getDy1());
        log.info("dx2 ：" + anchor_0.getDx2());
        log.info("dy2 ：" + anchor_0.getDy2());
        log.info("col1：" + anchor_0.getCol1());
        log.info("row1：" + anchor_0.getRow1());
        log.info("col2：" + anchor_0.getCol2());
        log.info("row2：" + anchor_0.getRow2());
        if (createFlag) {
            Picture picture_0 = patriarch.createPicture(anchor_0, workbook.addPicture(byteArrayOut_0.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        }
    }


    public static void buildExcelImage4(String imagePath, Sheet sheet, Drawing patriarch, Workbook workbook, int firstRow, int firstCol) throws IOException {
        ByteArrayOutputStream bytearrayout = new ByteArrayOutputStream();
        XSSFClientAnchor xssfClientAnchor = new XSSFClientAnchor();
        xssfClientAnchor.setAnchorType(XSSFClientAnchor.MOVE_AND_RESIZE);//移动
        log.info("图片路径：" + imagePath);
        log.info("(" + firstRow + "," + firstCol + ")");
        File image = new File(imagePath);
        BufferedImage userHeadImg = DrawImageUtils.drawImage(image);
        ImageIO.write(userHeadImg, "jpg", bytearrayout);


        //======================================获取图片属性  begin====================================
        int imgHeightPixel = userHeadImg.getHeight();// 图片高度
        int imgWidthPixel = userHeadImg.getWidth();// 图片宽度
        //======================================获取图片属性  end====================================

        // 获取合并单元格
        CellRangeAddress mergedRegion = ExcelUtils.getMergedRegion(sheet, firstRow, firstCol);

        //======================================循环计算 合并单元格 高度和宽度  begin====================================
        int cellHeightPixel = 0;
        for (int row = mergedRegion.getFirstRow(); row <= mergedRegion.getLastRow(); row++) {
            // getHeightInPoints()方法获取的是点（磅），就是excel设置的行高，1英寸有72磅，一般显示屏一英寸是96个像素
            cellHeightPixel += sheet.getRow(row).getHeightInPoints() / 72 * 96;
        }
        double cellWidthPixel = 0;
        for (int col = mergedRegion.getFirstColumn(); col <= mergedRegion.getLastColumn(); col++) {
            cellWidthPixel += sheet.getColumnWidthInPixels(col);
        }
        //======================================循环计算 合并单元格 高度和宽度  end====================================


        double widthRatio = imgWidthPixel / cellWidthPixel;//宽
        double heightRatio = imgHeightPixel / cellHeightPixel;//高
        double ratio = 1;
        double needWidthPixel = 0;// 计算需要的宽高 单位毫米
        double needHeightPixel = 0;

        if (imgWidthPixel > cellWidthPixel && imgHeightPixel > cellHeightPixel) {
            // 图片的宽和高  均大于单元格
            log.info("图片的宽和高均大于单元格");
            ratio = Math.max(widthRatio, heightRatio);
            needWidthPixel = Math.abs(imgWidthPixel / ratio);// 计算需要的宽高
            needHeightPixel = Math.abs(imgHeightPixel / ratio);
        } else if (imgWidthPixel < cellWidthPixel && imgHeightPixel < cellHeightPixel) {
            // 图片的宽和高  均小于单元格
            log.info("图片的宽和高均小于单元格");
            ratio = Math.max(widthRatio, heightRatio);
            needWidthPixel = Math.abs(imgWidthPixel / ratio);
            needHeightPixel = Math.abs(imgHeightPixel / ratio);
        } else if (imgWidthPixel > cellWidthPixel && imgHeightPixel < cellHeightPixel) {
            // 图片的宽大于单元格的宽,且图片的高小于单元格的高
            log.info("图片的宽大于单元格的宽,且图片的高小于单元格的高");
            ratio = heightRatio;
            needWidthPixel = Math.abs(imgWidthPixel / ratio);
            needHeightPixel = cellHeightPixel;
        } else if (imgWidthPixel < cellWidthPixel && imgHeightPixel > cellHeightPixel) {
            // 图片的宽小于单元格的宽,且图片的高大于单元格的高
            log.info("图片的宽小于单元格的宽,且图片的高大于单元格的高");
            ratio = widthRatio;
            needWidthPixel = cellWidthPixel;
            needHeightPixel = Math.abs(imgHeightPixel / ratio);
        }


        // 计算宽所在的列
        int needColNum = 0;
        double hasWidthPixel = 0D;
        for (int col = mergedRegion.getFirstColumn(); col <= mergedRegion.getLastColumn(); col++) {
            if (hasWidthPixel >= needWidthPixel) {
                break;
            }
            hasWidthPixel += sheet.getColumnWidthInPixels(col);
            needColNum++;
        }

        // 计算横轴空白
        double spaceWidthPixels = hasWidthPixel - needWidthPixel;
        double colCoordinatesPerPixels = 0.0D;
        // 获得宽所在的列
        float columnWidthPixel = sheet.getColumnWidthInPixels(mergedRegion.getFirstColumn() + needColNum - 1);
        colCoordinatesPerPixels = ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS / columnWidthPixel;
        int pictureWidthCoordinates = (int) (spaceWidthPixels * colCoordinatesPerPixels);


        //  =====================

        // 计算纵轴空白
        int needRowNum = 0;
        double hasHeightPixel = 0D;
        for (int row = mergedRegion.getFirstRow(); row <= mergedRegion.getLastRow(); row++) {
            if (hasHeightPixel >= needHeightPixel) {
                break;
            }
            hasHeightPixel += sheet.getRow(row).getHeightInPoints() / 72 * 96;
            needRowNum++;
        }

        double spaceHeightMM = hasHeightPixel - needHeightPixel;//计算空白
        double rowCoordinatesPerPixel = 0.0D;
        double rowHeightPixels = sheet.getRow(mergedRegion.getFirstRow() + needRowNum - 1).getHeightInPoints() / 72 * 96;
        //每个单元格宽度分成1023份  高分成256份
        rowCoordinatesPerPixel = ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS / rowHeightPixels;
        int pictureHeightCoordinates = (int) (spaceHeightMM * rowCoordinatesPerPixel);

        int spacePixel = 0;// 四周留白
        int dx1 = spacePixel;
        int dy1 = spacePixel;
        int dx2 = ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS - pictureWidthCoordinates + spacePixel;
        int dy2 = ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - pictureHeightCoordinates + spacePixel;

        //int dx2 = (ExcelUtils.TOTAL_COLUMN_COORDINATE_POSITIONS);
        //int dy2 = (ExcelUtils.TOTAL_ROW_COORDINATE_POSITIONS - 26);

        int col1 = mergedRegion.getFirstColumn();
        int row1 = mergedRegion.getFirstRow();
        int col2 = mergedRegion.getFirstColumn() + needColNum - 1;
        int row2 = mergedRegion.getFirstRow() + needRowNum - 1;

        xssfClientAnchor.setDx1(Math.round(dx1 * XSSFShape.EMU_PER_PIXEL));
        xssfClientAnchor.setDy1(Math.round(dy1 * XSSFShape.EMU_PER_PIXEL));
        xssfClientAnchor.setDx2(Math.round(dx2 * XSSFShape.EMU_PER_PIXEL));
        xssfClientAnchor.setDy2(Math.round(dy2 * XSSFShape.EMU_PER_PIXEL));

        xssfClientAnchor.setCol1(col1);
        xssfClientAnchor.setRow1(row1);
        xssfClientAnchor.setCol2(col2);
        xssfClientAnchor.setRow2(row2);

        log.info("widthRatio: " + widthRatio);
        log.info("heightRatio: " + heightRatio);
        log.info("ratio: " + ratio);

        log.info("imgWidthPixel:" + imgWidthPixel);
        log.info("cellWidthPixel:" + cellWidthPixel);
        log.info("needWidthPixel:" + needWidthPixel);
        log.info("hasWidthPixel:" + hasWidthPixel);

        log.info("imgHeightPixel:" + imgHeightPixel);
        log.info("cellHeightPixel:" + cellHeightPixel);
        log.info("needHeightPixel:" + needHeightPixel);
        log.info("hasHeightPixel:" + hasHeightPixel);

        log.info("dx1 ：" + dx1);
        log.info("dy1 ：" + dy1);
        log.info("dx2 ：" + dx2);
        log.info("dy2 ：" + dy2);
        log.info("col1：" + xssfClientAnchor.getCol1());
        log.info("row1：" + xssfClientAnchor.getRow1());
        log.info("col2：" + xssfClientAnchor.getCol2());
        log.info("row2：" + xssfClientAnchor.getRow2());
        Picture picture_0 = patriarch.createPicture(xssfClientAnchor, workbook.addPicture(bytearrayout.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
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
