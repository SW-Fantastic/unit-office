package org.swdc.offices;


/**
 * Row迭代接口，用于循环生成Excel的一行数据
 * @param <E> Data Element
 * @param <C> Excel Cell
 */
public interface RowIteratorFunction<E,C> {

    /**
     *
     * @param cell Excel的cell，请以它为起点开始本行数据的生成
     * @param element 当前的数据对象，类型取决于你提供的对象的类型。
     */
    void accept(C cell, E element);

}
