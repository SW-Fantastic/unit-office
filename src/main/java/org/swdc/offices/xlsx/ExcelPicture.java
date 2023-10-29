package org.swdc.offices.xlsx;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Files;

/**
 * 适用于XLSX（XSSF）格式的图片创建器
 * @param <T> 创建本对象的类型
 */
public class ExcelPicture<T> {

    private ClientAnchor anchor;

    private XSSFSheet sheet;

    private T parent;

    public ExcelPicture(XSSFSheet sheet, T from) {
        this.sheet = sheet;
        XSSFWorkbook workbook = sheet.getWorkbook();
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
    public ExcelPicture<T> position(int rowBegin, int colBegin, int rowEnd, int colEnd) {
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
    public ExcelPicture<T> position(int row, int col) {
        return position(row,col,row,col);
    }

    /**
     * 跨行跨列处理
     * @param rowSpan 图片的纵向跨行数
     * @param colSpan 图片的横向跨列数
     * @return 本对象
     */
    public ExcelPicture<T> cross(int rowSpan, int colSpan) {
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
     * @param type 图片类型
     * @return 本对象
     */
    public ExcelPicture<T> file(File file, int type) {
        try {
            byte[] data = Files.readAllBytes(file.toPath());
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            int pictureIndex = sheet.getWorkbook().addPicture(data,type);
            XSSFPicture picture = drawing.createPicture(anchor,pictureIndex);
            picture.resize();
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
