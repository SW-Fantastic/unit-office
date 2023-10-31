package org.swdc.offices.generator;

/**
 * 阶段生成函数。
 *
 * 如果你在生成一个复杂的Excel，直接把所有的Excel API
 * 放在一个代码块中，这是非常不明智的，他不利于未来的修改和维护，
 * 使用本项目创建Excel，有时代码会十分冗长，通过本类可以协助你创建
 * 一个具备更高可维护性的Excel导出类。
 *
 * @param <S> 本函数正在导出的Excel表类型，应为ExcelSheet（XLSX/XSSF）或者
 *           ExcelXLSSheet（XLS/HSSF）
 */
public interface GenerateFunction<S> {

    void generate(PipedGenerationContext context,S sheet);

}
