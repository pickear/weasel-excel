package com.weasel.excel;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Sheet;
import jxl.Workbook;

import com.weasel.excel.exception.CreateWorkBookException;

/**
 * @author Dylan
 * @time 2013-5-14 下午5:08:36
 */
public class ExcelReader {

	private Workbook excel;
	
	public ExcelReader(InputStream inputStream){
		try {
			this.excel = Workbook.getWorkbook(inputStream);
		} catch (Exception e) {
			throw new CreateWorkBookException(e.getMessage());
		} 
	}
	
	/**
	 * @return
	 */
	public List<List<String>> read(){
		
		Sheet[] sheets = excel.getSheets();
		List<List<String>> result = new ArrayList<List<String>>();
		for (Sheet sheet : sheets) {
			int rows = sheet.getRows();
			int columns = sheet.getColumns();
			for(int i = 1;i < rows;i++){
				List<String> list = new ArrayList<String>();
				for(int j = 0;j < columns;j++){
					String content = sheet.getCell(j, i).getContents();
					list.add(content);
				}
				result.add(list);
			}
		}
		
		return result;
	}
}
