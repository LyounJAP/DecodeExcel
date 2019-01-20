package top.lyoun.excel;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 该类用来解析小数据量的execel，采用用户模式
 * 需要用到如下jar包
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
public class DecodeExcel {
	//文件路径
	public static final String PATH = System.getProperty("user.dir") + File.separator + "file" + File.separator;
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		String fileName = "文件名.xlsx"; //要解析的文件的文件名
		File file = new File(PATH + fileName);
		if (file.exists()) {
			InputStream inputStream = new FileInputStream(file);
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0);
			// 获得总列数
			int colNum = sheet.getRow(0).getPhysicalNumberOfCells();
			int rowNum = sheet.getLastRowNum();// 获得总行数
			System.out.println("总列数：" + colNum + " ,总行数：" + rowNum);
			DataFormatter formatter = new DataFormatter();
			StringBuffer sb = new StringBuffer();
			for (int j = 1; j < rowNum; j++) {
					Row row = sheet.getRow(j);
					for (Cell cell : row) {
						String text = formatter.formatCellValue(cell);
						System.out.print(text + ",");
						sb.append(text).append(",");
					}
					System.out.println();
			}
		} else{
			System.out.println("文件不存在");
		}
	}

	// 把ANSI编码的文件转换为utf-8编码文件
	public static File ansiToUtf8(String fileName) throws IOException {
		fileName = PATH + fileName;
		File srcFile = new File(fileName);
		File targetFile = new File(PATH + "sp.csv");
		if (!srcFile.exists()) {
			System.out.println("源文件不存在");
		} else {// 源文件存在
			if (!targetFile.exists()) {
				targetFile.createNewFile();
			}
			InputStreamReader isr = new InputStreamReader(new FileInputStream(srcFile), "GBK"); // ANSI编码
			BufferedReader br = new BufferedReader(isr);
			OutputStreamWriter osw = new OutputStreamWriter(new FileOutputStream(targetFile), "UTF-8"); // 存为UTF-8
			String line = br.readLine();
			while (null != line) {
				osw.write(line);
				osw.write(System.getProperty("line.separator"));
				line = br.readLine();
			}
			// 刷新缓冲区的数据，强制写入目标文件
			osw.flush();
			osw.close();
			isr.close();
			isr = new InputStreamReader(new FileInputStream(targetFile), "utf-8");
			br = new BufferedReader(isr);
			osw = new OutputStreamWriter(new FileOutputStream(srcFile), "utf-8");
			line = br.readLine();
			while (null != line) {
				osw.write(line);
				osw.write(System.getProperty("line.separator"));
				line = br.readLine();
			}
			osw.flush();
			osw.close();
			isr.close();
			targetFile.delete();
		}
		return srcFile;
	}
}
