package org.swdc.offices.xlsx;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;

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

    public ExcelPicture<T> position(int rowBegin, int colBegin, int rowEnd, int colEnd) {
        anchor.setCol1(colBegin);
        anchor.setCol2(colEnd);
        anchor.setRow1(rowBegin);
        anchor.setRow2(rowEnd);
        return this;
    }

    public ExcelPicture<T> position(int row, int col) {
        return position(row,col,row,col);
    }

    public ExcelPicture<T> cross(int rowSpan, int colSpan) {
        return position(
                anchor.getRow1(),
                anchor.getCol1(),
                anchor.getRow1() + rowSpan,
                anchor.getCol1() + colSpan
        );
    }

    public ExcelPicture<T> file(File file, int type) {
        try {
            FileInputStream fin = new FileInputStream(file);
            byte[] data = fin.readAllBytes();
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            int pictureIndex = sheet.getWorkbook().addPicture(data,type);
            XSSFPicture picture = drawing.createPicture(anchor,pictureIndex);
            picture.resize();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return this;
    }

    public T back() {
        return parent;
    }

}
