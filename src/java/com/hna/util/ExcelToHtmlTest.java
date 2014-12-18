package com.hna.util;

import java.io.IOException;

public class ExcelToHtmlTest {

	public static void main(String[] args) throws Exception {
//	InputStream is = new FileInputStream("D:/GitHub/ExcelPoiUtil/ExcelPoiUtil/src/file/excel_1418661501653.html");
//		BufferedReader bReader = new BufferedReader(new InputStreamReader(is,"GBK"));
//		StringBuffer sb = new StringBuffer();
//		String line = new String();
//		while ((line = bReader.readLine()) != null) {
//			sb.append(line);
//		}
//		bReader.close();
//		is.close();
//		System.out.println(sb.toString());
		
		String path = "/home/hunter/Downloads/poi_test/海航酒店集团酒店运营情况综合日报.xls";
		try {
			
			  StringBuffer sb = new StringBuffer(); ExcelToHtml eth =
			  ExcelToHtml.create(path, sb); eth.setCompleteHTML(true);
			  eth.print();
			 
			ExcelToHtml.execute(path);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
