package com.weasel.excel.test;

/**
 * 
 * @author Dylan
 * @time 2013-5-14 下午4:37:40
 */
public class User {

	private String name;
	private String passwd;
	
	public String getName() {
		return name;
	}
	public User setName(String name) {
		this.name = name;
		return this;
	}
	public String getPasswd() {
		return passwd;
	}
	public User setPasswd(String passwd) {
		this.passwd = passwd;
		return this;
	}
}
