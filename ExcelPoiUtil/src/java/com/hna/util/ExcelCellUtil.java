package com.hna.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCellUtil {
public static void main(String[] args) {
	try {
		ExcelCellUtil.createWb("E:/海航酒店集团酒店运营情况综合日报.xls");
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
}
	private static Log log = LogFactory.getLog(ExcelCellUtil.class);
	public static final Workbook createWb(String filePath) throws IOException {
		if (StringUtils.isBlank(filePath)) {
			throw new IllegalArgumentException("参数错误!!!");
		}
		File file = new File(filePath);
		if (file.getName().trim().toLowerCase().endsWith(Constant.EXCEL_SUFFIX)) {
			return new HSSFWorkbook(new FileInputStream(file));
		} else if (file.getName().trim().toLowerCase().endsWith(Constant.EXCELX_SUFFIX)) {
			return new XSSFWorkbook(new FileInputStream(file));
		} else {
			throw new IllegalArgumentException("不支持除：xls/xlsx以外的文件格式!!!");
		}
	}

	public static final Sheet getSheet(Workbook wb, String sheetName) {
		return wb.getSheet(sheetName);
	}

	public static final Sheet getSheet(Workbook wb, int index) {
		return wb.getSheetAt(index);
	}

	/**
	 * 获取单元格内文本信息
	 * 
	 * @param cell
	 * @return
	 * @date 2013-5-8
	 */
	public static final Object getValueFromCell(Cell cell) {
		if (cell == null) {
			log.debug("Cell is null !!!");
			return null;
		}
		Object value = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // 数字
			if (HSSFDateUtil.isCellDateFormatted(cell)) { // 如果是日期类型
				value = new SimpleDateFormat("yyyy-MM-dd").format(cell
						.getDateCellValue());
			} else

				value = getRoundHalfUp(cell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING: // 字符串
			value = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_FORMULA: // 公式
			// 用数字方式获取公式结果，根据值判断是否为日期类型

			double numericValue = cell.getNumericCellValue();
			/*
			 * if(HSSFDateUtil.isValidExcelDate(numericValue)) { // 如果是日期类型
			 * value = new
			 * SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue()) ;
			 * } else
			 */
			value = getRoundHalfUp(numericValue);
			break;
		case Cell.CELL_TYPE_BLANK: // 空白
			// value = StringUtils.EMPTY;
			value = new Double(0);
			break;
		/*
		 * case Cell.CELL_TYPE_BOOLEAN: // Boolean value =
		 * String.valueOf(cell.getBooleanCellValue()); break;
		 */
		/*
		 * case Cell.CELL_TYPE_ERROR: // Error，返回错误码 value =
		 * String.valueOf(cell.getErrorCellValue()); break;
		 */
		default:
			// value = StringUtils.EMPTY;
			value = new Double(0);
			break;
		}
		// 使用[]记录坐标
		return value;// + "["+cell.getRowIndex()+","+cell.getColumnIndex()+"]" ;
	}

	/**
	 * @param cell
	 * @return
	 */
	public static final Cell getCloneCell(Cell cell) {
		//ToDo:封装获取的cell
		return cell;
	}
	
	/**
	 * @param sheet
	 * @return
	 */
	public static final List<Cell> getCellListFromSheet(Sheet sheet){
		//ToDo:从制定Sheet获取所需内容
		List<Cell> cellList = new ArrayList<Cell>();
		return cellList;
	} 


	public static Double getRoundHalfUp(double dobuleData) {
		return dobuleData;
		// return new BigDecimal(dobuleData).setScale(2,
		// BigDecimal.ROUND_HALF_UP).doubleValue();
	}
}
