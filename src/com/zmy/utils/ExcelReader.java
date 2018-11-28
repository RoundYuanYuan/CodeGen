package com.zmy.utils;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReader {
	public static void main(String[] args) throws Exception {
		InputStream inp = new FileInputStream("/home/user/work/MCUS/input/imicro_generator.xls");
		// InputStream inp = new FileInputStream("workbook.xlsx");
		Workbook wb = WorkbookFactory.create(inp);
		Sheet sheet = wb.getSheet("Structures");
		//遍历指定列
		int colINdex = ColUtils.colName2colIndex("Structure member name", 1, sheet);
		Iterator<Row> RowIter = sheet.iterator();
		while(RowIter.hasNext())
		{
			Row row = RowIter.next();
			Cell cell = row.getCell(colINdex);
			if(cell!=null) {
			String value = cell.getStringCellValue();
			System.out.println(value+" "+cell.getRowIndex());
			}
		}
		System.out.println("读取完毕");
		
		inp.close();
	}
}
