package com.zmy.utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ColUtils {
	/**
	 * @author RoundYuan
	 * @param  COL_Name 要转换的列名
	 * @param  COL_Name_Index 每个Sheet列名所在的行
	 * @param  sheet 目标Sheet
	 * @return COL_Index 返回列的实际位置
	 */
	public static int colName2colIndex(String COL_Name, int COL_Name_Index, Sheet sheet) {
		int COL_Index=-1;
		Row row = sheet.getRow(COL_Name_Index);
		Iterator<Cell> rowIter = row.iterator();
		while(rowIter.hasNext())
		{
			Cell currCell = rowIter.next();
			String currCOLName = currCell.getStringCellValue();
			if(currCOLName.equals(COL_Name))
				COL_Index=currCell.getColumnIndex();
		}
		return COL_Index;
	}
	
	
}
