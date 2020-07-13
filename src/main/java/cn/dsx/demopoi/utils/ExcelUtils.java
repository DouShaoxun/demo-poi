package cn.dsx.demopoi.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;

/**
 * @Classname: ExcelUtils
 * @Author: Dsx
 * @Date: 2020/07/12/19:16
 */
public class ExcelUtils {
    /**
     * 获取合并单元格的值
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if (row >= firstRow && row <= lastRow) {

                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);

                    return getCellValue(fCell);
                }
            }
        }

        return null;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 获取单元格的值
     *
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell) {

        if (cell == null) {
            return "";
        }

        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

            return cell.getStringCellValue();

        } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {

            return String.valueOf(cell.getBooleanCellValue());

        } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

            return cell.getCellFormula();

        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {

            return String.valueOf(cell.getNumericCellValue());

        }

        return "";
    }

    /**
     * 获取单元格起始位置
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static CellRangeAddress getMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return ca;
                }
            }
        }
        return new CellRangeAddress(row, row, column, column);
    }
    //1023和255指的是每个单元格被切分的份数，指定的是最后的单元格的最右下角的一个点
    public static final int TOTAL_COLUMN_COORDINATE_POSITIONS = 1023; // MB
    public static final int TOTAL_ROW_COORDINATE_POSITIONS = 255;     // MB

    public static class ConvertImageUnits {

        // Each cell contains a fixed number of co-ordinate points; this number
        // does not vary with row height or column width or with font. These two
        // constants are defined below.

        // The resolution of an image can be expressed as a specific number
        // of pixels per inch. Displays and printers differ but 96 pixels per
        // inch is an acceptable standard to belong with.
        public static final int PIXELS_PER_INCH = 96;                     // MB
        // Constants that defines how many pixels and points there are in a
        // millimetre. These values are required for the conversion algorithm.
        public static final double PIXELS_PER_MILLIMETRES = 3.78;         // MB
        public static final double POINTS_PER_MILLIMETRE = 2.83;          // MB
        // The column width returned by HSSF and the width of a picture when
        // positioned to exactly cover one cell are different by almost exactly
        // 2mm - give or take rounding errors. This constant allows that
        // additional amount to be accounted for when calculating how many
        // cells the image ought to overlie.
        public static final double CELL_BORDER_WIDTH_MILLIMETRES = 2.0D;  // MB
        public static final short EXCEL_COLUMN_WIDTH_FACTOR = 256;
        public static final int UNIT_OFFSET_LENGTH = 7;
        public static final int[] UNIT_OFFSET_MAP = new int[]
                {0, 36, 73, 109, 146, 182, 219};

        /**
         * pixel units to excel width units(units of 1/256th of a character width)
         *
         * @param pxs
         * @return
         */
        public static short pixel2WidthUnits(int pxs) {
            short widthUnits = (short) (EXCEL_COLUMN_WIDTH_FACTOR *
                    (pxs / UNIT_OFFSET_LENGTH));
            widthUnits += UNIT_OFFSET_MAP[(pxs % UNIT_OFFSET_LENGTH)];
            return widthUnits;
        }

        /**
         * excel width units(units of 1/256th of a character width) to pixel
         * units.
         *
         * @param widthUnits
         * @return
         */
        public static int widthUnits2Pixel(short widthUnits) {
            int pixels = (widthUnits / EXCEL_COLUMN_WIDTH_FACTOR)
                    * UNIT_OFFSET_LENGTH;
            int offsetWidthUnits = widthUnits % EXCEL_COLUMN_WIDTH_FACTOR;
            pixels += Math.round(offsetWidthUnits /
                    ((float) EXCEL_COLUMN_WIDTH_FACTOR / UNIT_OFFSET_LENGTH));
            return pixels;
        }

        /**
         * Convert Excels width units into millimetres.
         *
         * @param widthUnits The width of the column or the height of the
         *                   row in Excels units.
         * @return A primitive double that contains the columns width or rows
         * height in millimetres.
         */
        public static double widthUnits2Millimetres(short widthUnits) {
            return (ConvertImageUnits.widthUnits2Pixel(widthUnits) /
                    ConvertImageUnits.PIXELS_PER_MILLIMETRES);
        }

        /**
         * Convert into millimetres Excels width units..
         *
         * @param millimetres A primitive double that contains the columns
         *                    width or rows height in millimetres.
         * @return A primitive int that contains the columns width or rows
         * height in Excels units.
         */
        public static int millimetres2WidthUnits(double millimetres) {
            return (ConvertImageUnits.pixel2WidthUnits((int) (millimetres *
                    ConvertImageUnits.PIXELS_PER_MILLIMETRES)));
        }

        public static int pointsToPixels(double points) {
            return (int) Math.round(points / 72D * PIXELS_PER_INCH);
        }

        public static double pointsToMillimeters(double points) {
            return points / 72D * 25.4;
        }
    }


    //参考 http://cn.voidcc.com/question/p-wotldxzk-ot.html
    // set padding between picture and gridlines so gridlines would not covered by the picture
    private static final double paddingSize = 2;
    private static final int padding = Units.toEMU(paddingSize);

    public static int[] calCellAnchor(double cellX, double cellY, int imgX, int imgY) {
        // assume Y has fixed padding first
        return calCoordinate(true, cellX, cellY, imgX, imgY);
        //return calCoordinate(true,  cellY, cellX,imgX, imgY);
    }

    public static int[] calCoordinate(boolean fixTop, double cellX, double cellY, int imgX, int imgY) {

        double ratioImg = ((double) imgX) / imgY;// 图片比例
        double ratioCell = ((double) cellX) / cellY;// 单元格比例
        int x = imgX;
        System.out.println("ratioImg:"+ratioImg);
        System.out.println("ratioCell:"+ratioCell);
        // 2 * paddingSize是两侧留白
        //if (ratioImg > ratioCell) {
        //    x = imgY;
        //    x = (int) Math.round(Units.toEMU(cellX - 2 * paddingSize) * ratioImg);
        //    x = (Units.toEMU(cellY) - x) / 2;
        //
        //}else{
        //    x = imgX;
        //    x = (int) Math.round(Units.toEMU(cellY - 2 * paddingSize) * ratioImg);
        //    x = (Units.toEMU(cellX) - x) / 2;
        //}

        x = (int) Math.round(Units.toEMU(cellY - 2 * paddingSize) * ratioImg);
        x = (Units.toEMU(cellX) - x)/2;

        if (x < padding) {
            return calCoordinate(false, cellY, cellX, imgY, imgX);
        }
        return calDirection(fixTop, x);
    }

    public static int[] calDirection(boolean fixTop, int x) {
        if (fixTop) {
            return new int[]{x, padding, -x, -padding};
        } else {
            return new int[]{padding, x, -padding, -x};
        }
    }

}
