package com.hna.util;

import java.io.IOException;

public class ExcelToHtmlTest {

	 public static void main(String[] args) throws Exception{
		 
		 String path = "/home/hunter/Downloads/poi_test/海航酒店集团酒店运营情况综合日报.xls";
		 try {
	/*		 StringBuffer sb = new StringBuffer();
			 ExcelToHtml eth = ExcelToHtml.create(path, sb);
			 eth.setCompleteHTML(true);
			 eth.print();*/
			 ExcelToHtml.execute(path);
		} catch (IOException e) {
			e.printStackTrace();
		}
	 }
	 
}
