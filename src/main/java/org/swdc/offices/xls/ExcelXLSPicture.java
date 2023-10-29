package org.swdc.offices.xls;

import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;


/**
 * 适用于XLS（HSSF）格式的图片创建器
 * @param <T> 创建本对象的类型
 */
public class ExcelXLSPicture<T> {

    private ClientAnchor anchor;

    private HSSFSheet sheet;

    private T parent;

    public ExcelXLSPicture(HSSFSheet sheet, T from) {
        this.sheet = sheet;
        HSSFWorkbook workbook = sheet.getWorkbook();
        CreationHelper helper = workbook.getCreationHelper();
        anchor = helper.createClientAnchor();
        this.parent = from;
    }

    /**
     * 指定图片的位置
     * @param rowBegin 图片将会从此行开始
     * @param colBegin 图片将会从此列开始
     * @param rowEnd 图片将会在此行结束
     * @param colEnd 图片将会在此列结束
     * @return 本对象
     */
    public ExcelXLSPicture<T> position(int rowBegin, int colBegin, int rowEnd, int colEnd) {
        anchor.setCol1(colBegin);
        anchor.setCol2(colEnd);
        anchor.setRow1(rowBegin);
        anchor.setRow2(rowEnd);
        return this;
    }

    /**
     * 修改图像插入的单元格位置
     * @param row 图片将会出现在此行
     * @param col 图片将会出现在此列
     * @return 本对象。
     */
    public ExcelXLSPicture<T> position(int row, int col) {
        return position(row,col,row,col);
    }

    /**
     * 跨行跨列处理
     * @param rowSpan 图片的纵向跨行数
     * @param colSpan 图片的横向跨列数
     * @return 本对象
     */
    public ExcelXLSPicture<T> cross(int rowSpan, int colSpan) {
        return position(
                anchor.getRow1(),
                anchor.getCol1(),
                anchor.getRow1() + rowSpan,
                anchor.getCol1() + colSpan
        );
    }

    /**
     * 图像的内容，需要一个File
     * @param file 图片文件
     * @return 本对象
     */
    public ExcelXLSPicture<T> file(File file) {
        try {
            BufferedImage image = ImageIO.read(file);
            // convert any format to png
            BufferedImage target = new BufferedImage(image.getWidth(),image.getHeight(),BufferedImage.TYPE_INT_ARGB);
            Graphics2D g2d = target.createGraphics();
            g2d.drawImage(image,0,0, image.getWidth(),image.getHeight(),Color.WHITE,null);

            ByteArrayOutputStream bot = new ByteArrayOutputStream();
            ImageIO.write(target,"png",bot);
            byte[] data = bot.toByteArray();
            int pictureIndex = sheet.getWorkbook().addPicture(data,HSSFWorkbook.PICTURE_TYPE_PNG);
            HSSFPatriarch drawing = sheet.createDrawingPatriarch();
            drawing.createPicture(anchor,pictureIndex);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return this;
    }

    /**
     * 结束插入，返回上层
     * @return 创建本对象的对象。
     */
    public T back() {
        return parent;
    }

}
