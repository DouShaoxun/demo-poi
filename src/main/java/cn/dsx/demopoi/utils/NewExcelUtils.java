package cn.dsx.demopoi.utils;

import java.text.SimpleDateFormat;

/**
 * @Classname: NewExcelUtils
 * @Author: Dsx
 * @Date: 2020/07/13/22:30
 */
public class NewExcelUtils {
    private static SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
    ///**
    // * 生成Excel
    // * @param targetPath 原excel文件地址
    // * @param imgPath 图片地址
    // * @param newFilePath 生成的新excel文件地址
    // * @param width 生成的新的图片宽度，单位是像素
    // */
    //public static void dataExportExcel(String targetPath, String imgPath, String newFilePath, double width) {
    //    InputStream input = null;
    //    OutputStream output = null;
    //    Workbook workbook = null;
    //    try {
    //        input = new FileInputStream(targetPath);
    //        if (!input.markSupported()) {
    //            input = new PushbackInputStream(input, 8);
    //        }
    //        if (POIFSFileSystem.hasPOIFSHeader(input)) {
    //            workbook= new HSSFWorkbook(input);// excel 2003
    //        } else if (POIXMLDocument.hasOOXMLHeader(input)) {
    //            workbook = new XSSFWorkbook(OPCPackage.open(input));// excel 2007
    //        }
    //        Sheet sheet = workbook.getSheetAt(0);
    //        for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
    //            Row row = sheet.getRow(i);
    //            if (row == null) {
    //                continue;
    //            }
    //            for (int j = 0; j < row.getLastCellNum(); j++) {
    //                Cell cell = row.getCell(j);
    //                if (cell == null) {
    //                    continue;
    //                }
    //                Comment comment = cell.getCellComment();
    //                if (comment == null) {
    //                    continue;
    //                }
    //                String commentValue = comment.getString().getString().trim();
    //                if (commentValue.equals("img")) {
    //                    NewExcelUtils.replaceImage(workbook, sheet, cell, imgPath, j, i, width);
    //                }
    //                cell.removeCellComment();
    //            }
    //        }
    //        File file = new File(newFilePath);
    //        if (!file.exists()) {
    //            file.createNewFile();
    //        }
    //        output = new FileOutputStream(newFilePath);
    //        workbook.write(output);
    //    } catch (Exception e) {
    //        e.printStackTrace();
    //    } finally {
    //        if (input != null) {
    //            try {
    //                input.close();
    //            } catch (IOException e) { }
    //        }
    //        if (output != null) {
    //            try {
    //                output.close();
    //            } catch (IOException e) { }
    //        }
    //    }
    //}
    //
    ///**
    // * 替换图片
    // * @param book
    // * @param sheet
    // * @param cell
    // * @param jdImagePath
    // * @param jdcol
    // * @param jdrow
    // * @param width
    // * @return
    // * @throws Exception
    // */
    //public static Workbook replaceImage(Workbook book, Sheet sheet, Cell cell, String jdImagePath, int jdcol, int jdrow, double width) throws Exception {
    //    InputStream jdis;
    //    byte[] jdbytes = null;
    //    try {
    //        jdis = new FileInputStream(jdImagePath);
    //        jdbytes = IOUtils.toByteArray(jdis);
    //    } catch (Exception e) {
    //        e.printStackTrace();
    //    }
    //    CreationHelper helper = book.getCreationHelper();
    //    Drawing drawing  = null;
    //    if (sheet instanceof XSSFSheet) {
    //        XSSFSheet xSSFSheet = (XSSFSheet)sheet;
    //        drawing = xSSFSheet.getDrawingPatriarch();
    //    } else if (sheet instanceof HSSFSheet) {
    //        HSSFSheet hSSFSheet = (HSSFSheet)sheet;
    //        drawing = hSSFSheet.getDrawingPatriarch();
    //    }
    //    if (drawing == null) {
    //        drawing = sheet.createDrawingPatriarch();
    //    }
    //    // 图片插入坐标
    //    if (-1 != jdcol && -1 != jdrow) {
    //        int jdpictureIdx = book.addPicture(jdbytes, Workbook.PICTURE_TYPE_JPEG);// 根据需要调整参数，如果是PNG，就改为 Workbook.PICTURE_TYPE_PNG
    //        ClientAnchor jdanchor = helper.createClientAnchor();
    //        jdanchor.setCol1(jdcol);
    //        jdanchor.setRow1(jdrow);
    //        // 获取原图片的宽度和高度，单位都是像素
    //        File image = new File(jdImagePath);
    //        BufferedImage sourceImg = ImageIO.read(image);
    //        double imageWidth = sourceImg.getWidth();
    //        double imageHeight = sourceImg.getHeight();
    //        // 获取单元格宽度和高度，单位都是像素
    //        double cellWidth = sheet.getColumnWidthInPixels(cell.getColumnIndex());
    //        double cellHeight = cell.getRow().getHeightInPoints() / 72 * 96;// getHeightInPoints()方法获取的是点（磅），就是excel设置的行高，1英寸有72磅，一般显示屏一英寸是96个像素
    //        // 插入图片，如果原图宽度大于最终要求的图片宽度，就按比例缩小，否则展示原图
    //        Picture pict = drawing.createPicture(jdanchor, jdpictureIdx);
    //        if (imageWidth > width) {
    //            double scaleX = width / cellWidth;// 最终图片大小与单元格宽度的比例
    //            // 最终图片大小与单元格高度的比例
    //            // 说一下这个比例的计算方式吧：( imageHeight / imageWidth ) 是原图高于宽的比值，则 ( width * ( imageHeight / imageWidth ) ) 就是最终图片高的比值，
    //            // 那 ( width * ( imageHeight / imageWidth ) ) / cellHeight 就是所需比例了
    //            double scaleY = ( width * ( imageHeight / imageWidth ) ) / cellHeight;
    //            pict.resize(scaleX, scaleY);
    //        } else {
    //            pict.resize();
    //        }
    //    }
    //    return book;
    //}

    //public static void main(String[] args) throws IOException {
    //
    //    // 模板路径
    //    File directory = new File("src/main/resources");
    //    String courseFile = directory.getCanonicalPath();
    //    String excelPath = courseFile + "/templates/excel";
    //    String filePath = excelPath + "/template.xlsx";
    //    String imagePath = courseFile + "/static/image";
    //
    //    String exprotPath = excelPath + "/exprot/";
    //    File dir = new File(exprotPath);
    //    if (!dir.exists()) {
    //        dir.mkdir();
    //    }
    //    String format = simpleDateFormat.format(new Date());
    //    String outputFilePath = exprotPath + format + ".xlsx";
    //    NewExcelUtils.dataExportExcel(filePath, imagePath + "/1.jpg", outputFilePath, 70);
    //}
}
