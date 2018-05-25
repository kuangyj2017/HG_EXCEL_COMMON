/**
 * Copyright (c) 2014-2017 HG R&D Team.  All Rights Reserved.
 * This software is published under the HG Solution Team.
 * License version 2.0, a copy of which has been included with this
 * distribution in the LICENSE.txt file.
 *
 * @File name:  HGExcellRowVO.java
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

import java.util.List;

/**
 * @author kuangyj
 *
 */
public class HGExcelRowVO {
	//行号   以0起始
	private int rowIndex;
	//单元格对象列表
	private List<HGExcelCellVO> cellList;
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
	 * @return the cellList
	 */
	public List<HGExcelCellVO> getCellList() {
		return cellList;
	}
	/**
	 * @param cellList the cellList to set
	 */
	public void setCellList(List<HGExcelCellVO> cellList) {
		this.cellList = cellList;
	}	
}
