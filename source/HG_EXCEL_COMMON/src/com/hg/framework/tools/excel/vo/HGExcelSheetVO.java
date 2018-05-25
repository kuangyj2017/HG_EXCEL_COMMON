/**
 * Copyright (c) 2014-2017 HG R&D Team.  All Rights Reserved.
 * This software is published under the HG Solution Team.
 * License version 2.0, a copy of which has been included with this
 * distribution in the LICENSE.txt file.
 *
 * @File name:  HGExcelSheetVO.java
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
public class HGExcelSheetVO {
	//Sheet名称
	private String sheetName; 
	//每一行的记录
	private List<HGExcelRowVO> rowList;
	/**
	 * @return the sheetName
	 */
	public String getSheetName() {
		return sheetName;
	}
	/**
	 * @param sheetName the sheetName to set
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	/**
	 * @return the rowList
	 */
	public List<HGExcelRowVO> getRowList() {
		return rowList;
	}
	/**
	 * @param rowList the rowList to set
	 */
	public void setRowList(List<HGExcelRowVO> rowList) {
		this.rowList = rowList;
	}	
}
