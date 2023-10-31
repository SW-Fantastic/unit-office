package org.swdc.offices.generator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.swdc.offices.RowIteratorFunction;
import org.swdc.offices.xls.ExcelXLSCell;
import org.swdc.offices.xls.ExcelXLSRow;
import org.swdc.offices.xls.ExcelXLSSheet;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Function;

/**
 * Excel生成器类，用来快速创建格式简单的Excel的文档，
 * 你可以轻松使用本生成器生成包含一个表头和多行数据的简单Excel文件。
 * 适用于XLS（HSSF）格式。
 */
public class ExcelXLSSimpleGenerator {


    private Map<Class, RowIteratorFunction> rowIterators = new ConcurrentHashMap<>();

    private Function<ExcelXLSSheet, ExcelXLSRow> initExcelFunction = null;

    private HSSFWorkbook workbook = null;


    /**
     * 初始化一个Generator
     */
    public ExcelXLSSimpleGenerator() {
        try {
            this.workbook = new HSSFWorkbook();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 指定表格初始化函数
     * @param theinitFunction 表格初始化函数，你可以在这里做一些除了生成数据行之外的事情，
     *                        作者建议使用方法引用（Method Reference）
     * @return 本对象
     */
    public ExcelXLSSimpleGenerator generateExcelStructure(Function<ExcelXLSSheet,ExcelXLSRow> theinitFunction) {
        this.initExcelFunction = theinitFunction;
        return this;
    }


    /**
     * 添加行生成策略
     * @param type 数据类型
     * @param function 行生成函数，作者建议使用方法引用（Method Reference）
     * @return 本对象
     * @param <E> 数据类型
     */
    public final <E> ExcelXLSSimpleGenerator strategy(Class<E> type, RowIteratorFunction<E, ExcelXLSCell> function) {
        if (type == null || function == null){
            throw new RuntimeException("any parameter can not be null");
        }
        if (rowIterators.containsKey(type)) {
            throw new RuntimeException("can not register function for type: " + type.getName() + ", the function is already exist.");
        }
        rowIterators.put(type,function);
        return this;
    }

    /**
     * 使用本方法创建Excel。
     *
     * @param sheetName 生成到此Sheet中
     * @param items 生成的数据列表
     * @param outputStream Excel数据输出流
     * @throws IOException IO异常
     */
    public void createExcel(String sheetName, List<? extends Object> items, OutputStream outputStream) throws IOException {

        ExcelXLSSheet theSheet = new ExcelXLSSheet(workbook,sheetName);
        ExcelXLSRow row = null;
        if (initExcelFunction != null) {
            row = initExcelFunction.apply(theSheet);
        } else {
            row = theSheet.rowAt(0);
        }

        for (Object item: items) {
            RowIteratorFunction<Object,ExcelXLSCell> gen = rowIterators.get(item.getClass());
            if (gen != null) {
                gen.accept(row.cell(0),item);
                row = row.nextRow();
            }
        }

        workbook.write(outputStream);
    }


}
