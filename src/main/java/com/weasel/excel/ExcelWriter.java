package com.weasel.excel;

import java.beans.PropertyDescriptor;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import com.weasel.excel.annotation.ValueConvert;
import com.weasel.excel.exception.CreateWorkBookException;
import com.weasel.excel.exception.ReflectException;

/**
 * 
 * @author Dylan
 * @time 2013-5-14 下午3:01:59
 */
public class ExcelWriter {

	private WritableWorkbook excel;
	private WritableSheet sheet;
	private Integer row = 0;

	public ExcelWriter(OutputStream outputStream) {
		initContext(outputStream, "");
	}

	public ExcelWriter(OutputStream outputStream, String sheetName) {
		initContext(outputStream, sheetName);
	}

	/**当需要在excel的头部写入公司图片时，用此方法
	 * @param entities   要导出的数据
	 * @param title    标题名称
	 * @param headContent  头部内容
	 * @param imgPath   图片路径
	 */
	public <T> void write(List<T> entities,String[] title,String headContent,String imgPath){
		
		createHead(headContent, imgPath);
		createTitle(title);
		createExcelData(entities);
	}
	/**
	 * @param entities  要导出的数据
	 * @param title 标题名称
	 */
	public <T> void export(List<T> entities, String[] title) {

		createTitle(title);
		createExcelData(entities);
		
	}
	
	/**
	 * @param entities
	 */
	private <T> void createExcelData(List<T> entities){
		try {
			for (T entity : entities) {
				Class<?> clazz = entity.getClass();
				Field[] fields = clazz.getDeclaredFields(); // 得到类的所有字段
				int index = 0;
				for (Field field : fields) {
					// 得到属性的注解
					ValueConvert convert = field.getAnnotation(ValueConvert.class);
					PropertyDescriptor pd;
					pd = new PropertyDescriptor(field.getName(), clazz);
					//如果打上了@ValueConvert注解，就调用注解上指明的方法来得到值
					Method method = null == convert ? pd.getReadMethod() : clazz.getMethod(convert.method());
					Object o = method.invoke(entity); // 执行方法，得到返回值
					String value = null == o ? "" : o.toString();
					Label label = new Label(index++, row, value);
					sheet.addCell(label);
				}
				++row;
			}
			excel.write();
		} catch (Exception e) {
			throw new ReflectException(e.getMessage());
		} finally {
			close();
		}
	}

	/**
	 * @param outputStream
	 * @param sheetName
	 */
	private void initContext(OutputStream outputStream, String sheetName) {
		try {
			this.excel = Workbook.createWorkbook(outputStream);
			this.sheet = excel.createSheet(sheetName, 0);
		} catch (IOException e) {
			throw new CreateWorkBookException(e.getMessage());
		}
	}

	/**
	 * @param title
	 */
	private void createTitle(String[] title) {
		for (int i = 0; i < title.length; i++) {
			try {
				Label label = new Label(i, row, title[i]);
				sheet.addCell(label);
				sheet.setColumnView(i, 20);
			} catch (Exception e) {
				throw new RuntimeException(e.getMessage());
			}
		}
		++row;
	}

	/**
	 * @param sheet
	 * @param headContent
	 * @param imgPath
	 */
	private void createHead(String headContent, String imgPath) {

		File img = new File(imgPath);
		// 设置标题格式
		WritableFont wfTitle = new jxl.write.WritableFont(WritableFont.ARIAL, 35, WritableFont.BOLD, true);
		WritableCellFormat wcfTitle = new WritableCellFormat(wfTitle);
		// 设置水平对齐方式
		try {
			wcfTitle.setAlignment(Alignment.CENTRE);
			// 设置垂直对齐方式
			wcfTitle.setVerticalAlignment(VerticalAlignment.CENTRE);
			// 设置是否自动换行
			wcfTitle.setWrap(true);
			sheet.mergeCells(0, 0, 1, 0);
			sheet.mergeCells(2, 0, 12, 0); // 2:第三列 0 ：第一行 ，12:合并到第13列，0：只合并一行
			Label titleImg = new Label(0, row, "", wcfTitle);
			sheet.addCell(titleImg);
			Label titleContent = new Label(2, row++, headContent, wcfTitle);
			sheet.addCell(titleContent);
			WritableImage image = new WritableImage(0, 0, 2, 3.7, img);
			sheet.addImage(image);
		} catch (WriteException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 
	 */
	public void close() {
		if (null != excel) {
			try {
				excel.close();
			} catch (WriteException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}
