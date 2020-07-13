package cn.dsx.demopoi.utils;

/**
 * @Classname: ExcelFormat
 * @Author: Dsx
 * @Date: 2020/07/12/22:14
 */
public enum ExcelFormat {

    xlsx("xlsx"),
    xls("xls");


    ExcelFormat(String format) {
        this.format = format;
    }

    private String format;


    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }
}
