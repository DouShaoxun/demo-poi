package cn.dsx.demopoi.utils;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

/**
 * @Classname: DrawImageUtils
 * @Author: Dsx
 * @Date: 2020/07/12/17:56
 */
public class DrawImageUtils {
    public static BufferedImage drawImage(File file) throws IOException {
        BufferedImage image = ImageIO.read(file);
        String fileName = file.getName();
        String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
        if ("jpg".equalsIgnoreCase(suffix) || "jpeg".equalsIgnoreCase(suffix)) {
            //重画一下，要么上传图片变色或报“Invalid argument to native writeImage”错误
            BufferedImage tag = new BufferedImage(image.getWidth(), image.getHeight(), BufferedImage.TYPE_INT_BGR);
            Graphics g = tag.getGraphics();
            g.drawImage(image, 0, 0, null); // 绘制缩小后的图
            g.dispose();
            image = tag;
        }
        return image;
    }
}
