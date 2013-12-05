package com.weasel.excel.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import com.weasel.excel.ExcelReader;
import com.weasel.excel.ExcelWriter;

/*
 * 
 */
public class ExcelTest {

	@Test
	public void exportTest() {
		try {
			OutputStream outputStream = new FileOutputStream(new File("D:\\test.xls"));
			ExcelWriter excel = new ExcelWriter(outputStream, "测试excel");
			User u1 = new User().setName("张三").setPasswd("zhanshang");
			User u2 = new User().setName("李四").setPasswd("lishi");
			List<User> entities = new ArrayList<User>();
			entities.add(u1);
			entities.add(u2);
			excel.export(entities, new String[]{"名字","密码"});
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
	@Test
	public void exportTest2() {
		try {
			OutputStream outputStream = new FileOutputStream(new File("D:\\test.xls"));
			ExcelWriter excel = new ExcelWriter(outputStream, "测试excel");
			User u1 = new User().setName("张三").setPasswd("zhanshang");
			User u2 = new User().setName("李四").setPasswd("lishi");
			List<User> entities = new ArrayList<User>();
			entities.add(u1);
			entities.add(u2);
			excel.write(entities, new String[]{"名字","密码"},"员工信息报表",this.getClass().getResource("").getPath()+"/user.png");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
	
	@Test
	public void uploadTest2(){
		
		try {
			InputStream inputStream = new FileInputStream(new File("D:\\test.xls"));
			ExcelReader reader = new ExcelReader(inputStream);
			List<List<String>> result = reader.read();
			System.out.println(result.size());
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

}
