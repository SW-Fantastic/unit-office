package org.swdc.offices.generator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.swdc.offices.xls.ExcelXLSSheet;
import org.swdc.offices.xlsx.ExcelSheet;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayDeque;
import java.util.Deque;
import java.util.List;

/**
 * 如果你在生成一个复杂的Excel，直接把所有的Excel API
 * 放在一个代码块中，这是非常不明智的，这不利于未来的修改和维护，
 * 使用本项目创建Excel，有时代码会十分冗长，通过本类可以协助你创建
 * 一个具备更高可维护性的Excel导出类。
 *
 * 首先，请按照GenerateFunction的方法generate(PipedGenerationContext context,S sheet)
 * 的格式编写生成Excel内容的方法，请自行决定应该以什么方式划分Excel生成的阶段，例如，你可以把它划分为
 * 生成表头，生成内容，生成结尾，生成边框四个部分，分别提供四个方法，这些方法的参数需要与GenerateFunction的
 * generate方法一致。
 *
 * 接下来，请创建一个方法来生成本类的对象，通过此对象，你将得到需要的Excel文件。
 *
 * <pre>
 *     public class PipedDemoGenerator {
 *
 *          // 生成阶段1 ： 为Excel添加表头
 *         public ExcelRow generateHeader(PipedGenerationContext ctx, ExcelXLSSheet sheet) {
 *             sheet.autoColumnWidth(0)
 *                     .autoColumnWidth(1)
 *                     .autoColumnWidth(2)
 *                     .autoColumnWidth(3);
 *
 *             CellPresetFunction<ExcelXLSCell> preset = cell -> cell
 *                     .font()
 *                     .bold(true)
 *                     .back()
 *                     .alignCenter();
 *
 *             return sheet.rowAt(0).presetCell(preset)
 *                     .cell(0).text("姓名")
 *                     .nextCell().text("年龄")
 *                     .nextCell().text("生日")
 *                     .nextCell().text("性别")
 *                     .backToRow();
 *         }
 *
 *          // 生成阶段2 为Excel添加具体的数据内容。
 *         public void generatePerson(PipedGenerationContext ctx, ExcelXLSSheet sheet) {
 *             sheet.rowAt(1).forOf(ctx.getGrouped(Person.class), (cell, person) -> {
 *                 cell.text(person.getName()).nextCell()
 *                         .text(person.getAge()).nextCell()
 *                         .text(person.getBirthDay()).nextCell()
 *                         .text(person.getGender());
 *             });
 *         }
 *
 *          // 创建导出器
 *         public PipedExcelXLSGenerator createGenerator() {
 *             return new PipedExcelXLSGenerator()
 *                     .generateStage(this::generateHeader)
 *                     .generateStage(this::generatePerson);
 *         }
 *
 *     }
 * </pre>
 *
 */
public class PipedExcelXLSGenerator {

    private HSSFWorkbook workbook = null;

    private Deque<GenerateFunction<ExcelXLSSheet>> generateFunctions = new ArrayDeque<>();
    private Deque<GenerateFunction<ExcelXLSSheet>> working = new ArrayDeque<>();

    public PipedExcelXLSGenerator() {
        try {
            workbook = new HSSFWorkbook();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 添加生成阶段
     * @param stage 生成阶段函数，作者建议使用Java8的方法引用来实现（Method Reference）
     * @return 本对象
     */
    public PipedExcelXLSGenerator generateStage(GenerateFunction<ExcelXLSSheet> stage) {
        if (generateFunctions.contains(stage)){
            return this;
        }
        generateFunctions.addLast(stage);
        return this;
    }

    /**
     * 生成Excel
     * @param sheet 一个表格的名称，生成的内容会在这里面
     * @param data 数据列表，不限制对象的类型，你需要通过PipedGenerationContext依照它们的Class获取。
     * @param outputStream 输出流 Excel会写入这里。
     * @throws IOException
     */
    public void createExcel(String sheet, List<? extends Object> data, OutputStream outputStream) throws IOException {

        ExcelXLSSheet theSheet = new ExcelXLSSheet(workbook,sheet);
        PipedGenerationContext context = new PipedGenerationContext(data);
        Deque<GenerateFunction<ExcelXLSSheet>> working = this.working;

        GenerateFunction<ExcelXLSSheet> currFunction = null;
        while ((currFunction = generateFunctions.pollFirst()) != null) {
            currFunction.generate(context,theSheet);
            working.addLast(currFunction);
        }

        this.working = generateFunctions;
        this.generateFunctions = working;

        this.workbook.write(outputStream);
    }

}
