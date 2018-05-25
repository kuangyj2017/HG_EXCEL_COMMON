/**
 * Copyright (c) 2014-2017 HG R&D Team.  All Rights Reserved.
 * This software is published under the HG Solution Team.
 * License version 2.0, a copy of which has been included with this
 * distribution in the LICENSE.txt file.
 *
 * @File name:  HGExcelCellVO.java
 * @Create on:  2018年5月8日
 * @Author   :  
 *
 * @ChangeList
 * ---------------------------------------------------
 * NO      Date               Editor             ChangeReasons
 * 1       2018年5月8日            kuangyj            Create 
 *
 */
package com.hg.framework.tools.excel.vo;

/**
 * 定义Excel的CELL单元对象
 * 
 * @author kuangyj
 *
 */
public class HGExcelCellVO {
	//行号 - 以0起始
	private int rowIndex;
	//列号 - 以0起始
	private int cellIndex;
	//单元格数据类型 --具体参见 Cell.TYPE
	private int cellType;
	//单元格数据内容 
	private String cellValue;
	/**
	 * @return the rowIndex
	 */
	public int getRowIndex() {
		return rowIndex;
	}
	/**
	 * @param rowIndex the rowIndex to set
	 */
	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}
	/**
	 * @return the cellIndex
	 */
	public int getCellIndex() {
		return cellIndex;
	}
	/**
	 * @param cellIndex the cellIndex to set
	 */
	public void setCellIndex(int cellIndex) {
		this.cellIndex = cellIndex;
	}
	/**
	 * @return the cellType
	 */
	public int getCellType() {
		return cellType;
	}
	/**
	 * @param cellType the cellType to set
	 */
	public void setCellType(int cellType) {
		this.cellType = cellType;
	}
	/**
	 * @return the cellValue
	 */
	public String getCellValue() {
		return cellValue;
	}
	/**
	 * @param cellValue the cellValue to set
	 */
	public void setCellValue(String cellValue) {
		this.cellValue = cellValue;
	}	
}
