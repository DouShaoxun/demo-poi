package cn.dsx.demopoi;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import javax.imageio.ImageIO;
import javax.swing.filechooser.FileSystemView;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.Date;


@SpringBootTest
class XLSPoiTest {

    @Test
    void contextLoads() throws Exception {
        poiTest();
    }

    void poiTest() throws IOException {
        // 获取桌面路径
        FileSystemView fsv = FileSystemView.getFileSystemView();
        String desktop = fsv.getHomeDirectory().getPath();
        String filePath = desktop + "/template.xlsx";

        File file = new File(filePath);
        OutputStream outputStream = new FileOutputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");
        XSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("订单号");
        row.createCell(2).setCellValue("下单时间");
        row.createCell(3).setCellValue("个数");
        row.createCell(4).setCellValue("单价");
        row.createCell(5).setCellValue("订单金额");
        row.setHeightInPoints(30); // 设置行的高度

        XSSFRow sheetRow1 = sheet.createRow(1);
        sheetRow1.createCell(0).setCellValue("1");
        sheetRow1.createCell(1).setCellValue("NO00001");

        // 日期格式化
        XSSFCellStyle cellStyle2 = workbook.createCellStyle();
        XSSFCreationHelper creationHelper = workbook.getCreationHelper();
        cellStyle2.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
        sheet.setColumnWidth(2, 20 * 256); // 设置列的宽度

        XSSFCell cell2 = sheetRow1.createCell(2);
        cell2.setCellStyle(cellStyle2);
        cell2.setCellValue(new Date());

        sheetRow1.createCell(3).setCellValue(2);


        // 保留两位小数
        XSSFCellStyle cellStyle3 = workbook.createCellStyle();
        cellStyle3.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        XSSFCell cell4 = sheetRow1.createCell(4);
        cell4.setCellStyle(cellStyle3);
        cell4.setCellValue(29.5);


        // 货币格式化
        XSSFCellStyle cellStyle4 = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontName("华文行楷");
        font.setFontHeightInPoints((short) 15);
        font.setColor(HSSFColor.RED.index);
        cellStyle4.setFont(font);

        XSSFCell cell5 = sheetRow1.createCell(5);
        cell5.setCellFormula("D2*E2");  // 设置计算公式

        // 获取计算公式的值
        XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(workbook);
        cell5 = e.evaluateInCell(cell5);
        System.out.println(cell5.getNumericCellValue());

        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
        BufferedImage user_headImg = ImageIO.read(new File("C:/Users/Dsx/Desktop/20200710225429.jpg"));
        ImageIO.write(user_headImg, "jpg", byteArrayOut);


        // sheet只能获取一个
        XSSFDrawing patriarch = sheet.createDrawingPatriarch();
        // 设置图片的属性
        int dx1 = 0;
        int dy1 = 0;
        int dx2 = 0;
        int dy2 = 0;
        int col1 = 1;
        int row1 = 11;
        int col2 = 5;
        int row2 = 20;

        XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
        //anchor.setAnchorType(DONT_MOVE_AND_RESIZE);
        // 插入图片 
        patriarch.createPicture(anchor, workbook.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
        workbook.setActiveSheet(0);
        workbook.write(outputStream);
        outputStream.close();
    }

}
