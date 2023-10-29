package org.swdc.offices;

import org.swdc.offices.xlsx.ExcelCell;

/**
 * 预设Cell的格式，
 * 你可以通过本方法为样式复杂的Cell提供一套预设，
 * 并且在通过preset方法应用到Cell中。
 */
public interface CellPresetFunction<T> {

    T accept(T cell);

}
