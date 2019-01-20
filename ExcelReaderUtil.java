package top.lyoun.excel.bysax;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import top.lyoun.bean.TotalReport;
import top.lyoun.util.OperateMysql;

/**
 * 调用ExcelXlsReader类和ExcelXlsxReader类对 excel2003和excel2007两个版本进行大批量数据读取
  需要用到如下jar包
 * poi-3.17.jar
 * poi-examples-3.17.jar
 * poi-excelant-3.17.jar
 * poi-ooxml-3.17.jar
 * poi-ooxml-schemas-3.17.jar
 * poi-scratchpad-3.17.jar
 * xerces-2.4.0.jar
 * xmlbeans-2.6.0.jar
 * commons-collections4-4.2.jar
 * @author Lyoun
 * 2019.01.20
 */
public class ExcelReaderUtil {
	// excel2003扩展名
	public static final String EXCEL03_EXTENSION = ".xls";
	// excel2007扩展名
	public static final String EXCEL07_EXTENSION = ".xlsx";
	public static final String PATH = System.getProperty("user.dir") + File.separator + "file" + File.separator;

	public static void main(String[] args) throws Exception {
		String fileName = "文件名 .xlsx";
		ExcelReaderUtil.readExcel(PATH + fileName);
		
	}

	/**
	 * 每获取一条记录，即打印 在flume里每获取一条记录即发送，而不必缓存起来，可以大大减少内存的消耗，这里主要是针对flume读取大数据量excel来说的
	 */
	public static void sendRows(String filePath, String sheetName, int sheetIndex, int curRow, List<String> cellList) {
		// 打印excel提取的数据
		printExecel(curRow, cellList); //curRow：当前行数，对应于excel中的行数，cellList，Excel返回的一行中所有列的数据集合
		//你可以在这个方法里面执行你自己的代码逻辑
	}

	// 该方法专门用来打印excel提取的数据
	public static void printExecel(int curRow, List<String> cellList) {
		StringBuffer oneLineSb = new StringBuffer();
		oneLineSb.append("row" + curRow);
		oneLineSb.append("::");
		int curCol = 0;
		for (String cell : cellList) {//循环遍历cellList数据
			oneLineSb.append(cell.trim());
			oneLineSb.append("|");
			curCol++;
		}
		oneLineSb.append("::");
		oneLineSb.append("sumCol" + curCol);
		String oneLine = oneLineSb.toString();
		if (oneLine.endsWith("|")) {
			oneLine = oneLine.substring(0, oneLine.lastIndexOf("|"));
		} // 去除最后一个分隔符
		System.out.println(oneLine);
	}

	public static void readExcel(String fileName) throws Exception {
		System.out.println("开始解析文件，请稍等...");
		int totalRows = 0;
		if (fileName.endsWith(EXCEL03_EXTENSION)) { // 处理excel2003文件
			ExcelXlsReader excelXls = new ExcelXlsReader();
			totalRows = excelXls.process(fileName);
		} else if (fileName.endsWith(EXCEL07_EXTENSION)) {// 处理excel2007文件
			ExcelXlsxReader excelXlsxReader = new ExcelXlsxReader();
			totalRows = excelXlsxReader.process(fileName);
		} else {
			throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
		}
		System.out.println("发送的总行数：" + totalRows);
	}
}
