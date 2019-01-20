package top.lyoun.excel.bysax;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import top.lyoun.bean.TotalReport;
import top.lyoun.util.OperateMysql;

/**
 * 调用ExcelXlsReader类和ExcelXlsxReader类对 excel2003和excel2007两个版本进行大批量数据读取
 */
public class ExcelReaderUtil {
	// excel2003扩展名
	public static final String EXCEL03_EXTENSION = ".xls";
	// excel2007扩展名
	public static final String EXCEL07_EXTENSION = ".xlsx";
	public static final String PATH = System.getProperty("user.dir") + File.separator + "file" + File.separator;

	public static void main(String[] args) throws Exception {
		String fileName = "11月报表 .xlsx";
		ExcelReaderUtil.readExcel(PATH + fileName);
		
	}

	/**
	 * 每获取一条记录，即打印 在flume里每获取一条记录即发送，而不必缓存起来，可以大大减少内存的消耗，这里主要是针对flume读取大数据量excel来说的
	 */
	public static void sendRows(String filePath, String sheetName, int sheetIndex, int curRow, List<String> cellList) {
		//if(curRow<=10)
		//printExecel(curRow, cellList); // 打印excel提取的数据
		try {
			if(curRow == 3) {createDicts(cellList);}//创建字典
			//else if(curRow > 3) {storageToMysql(cellList);}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} // 把execel提取到的数据保存到数据库中

	}

	//创建字典表并插入数据
	private static void createDicts(List<String> cellList) throws Exception{
		StringBuffer oneLineSb = new StringBuffer();
		int sumCol = 0;
		for (String cell : cellList) {
			oneLineSb.append(cell.trim());
			oneLineSb.append("|");
			sumCol++;
		}
		String oneLine = oneLineSb.toString();
		if (oneLine.endsWith("|")) {
			oneLine = oneLine.substring(0, oneLine.lastIndexOf("|"));
		} // 去除最后一个分隔符
		String str[] = oneLine.split("\\|");
		if (str.length >= sumCol) {
			// 创建表t_totalReportDict
			String tabName = "t_totalReportDict";
			String arrays[] = { "id MEDIUMINT UNSIGNED PRIMARY KEY AUTO_INCREMENT", "name VARCHAR(255)",
					"description VARCHAR(255)"};
			OperateMysql.createTable(tabName, arrays);
			OperateMysql.truncateTable(tabName);
			//构造字典
			String engs[] = {"city", "county", "worksheetNum", "worksheetTheme", "worksheetType", "operateType",
					"currentLink", "currentLinkHandler", "sendBegTime", "endTime", "originator", "installAddr",
					"installAddrType", "CRMorderNumber", "acceptanceBusinessHall", "receivingSalesPerson",
					"acceptSalesmanPhone", "customerName", "customerContact", "customerContactPhone",
					"broadbandAccount", "coverType", "orderAcceptStartingTime", "orderAcceptEndingTime",
					"orderAcceptBacklogMan", "orderAcceptHandler", "openingBookingClosingTime",
					"openingBookingBacklogMan", "openingBookingHandler", "onsiteOpeningStartingTime",
					"onsiteOpeningBacklogMan", "onsiteOpeningHandler", "satisfactSurveyStartingTime",
					"satisfactSurveyEndingTime", "satisfactSurveyBacklogMan", "satisfactSurveyHandler",
					"subordinateBranch", "daiweiCompany", "maintenanceTeam", "realConstructerUnit", "realConstructer",
					"activateAutomatically", "duration", "submitByPhone", "villageName", "STBLicensingParty", "STPId",
					"sendModel", "replacementPreAddress", "replacementAddress", "lightCatSN", "replacementPONID",
					"replacementPrePONID", "fileInTime", "appointInTime", "smallMicroBroadband", "appointmentTime",
					"daiWeiUnit", "note", "temporaryConstructer", "temporaryConstructerPhone",
					"orderReachingNetworkTime", "outConstructionsFilingTime", "villageId"};
			Map<String, Object> insetMap = new HashMap<String, Object>();
			for(int i=0; i < str.length; i++) {
				insetMap.put("name",engs[i]);
				insetMap.put("description",str[i]);
				OperateMysql.insert(tabName, insetMap);
			}
		}
	}

	private static boolean firstCreate = true;

	// 该方法专门把execel提取到的数据保存到数据库中
	public static void storageToMysql(List<String> cellList) throws Exception {
		StringBuffer oneLineSb = new StringBuffer();
		int sumCol = 0;
		for (String cell : cellList) {
			oneLineSb.append(cell.trim());
			oneLineSb.append("|");
			sumCol++;
		}
		String oneLine = oneLineSb.toString();
		if (oneLine.endsWith("|")) {
			oneLine = oneLine.substring(0, oneLine.lastIndexOf("|"));
		} // 去除最后一个分隔符
		String str[] = oneLine.split("\\|");
		int j = 0;
		if (str.length >= sumCol) {
			Integer villageId = null;
			try {
				villageId = Integer.parseInt(str[j + 63]);
			} catch (Exception e) {
				villageId = null;
			}

			// 创建表t_totalReport
			String tabName = "t_totalReport";
			if (firstCreate) {
				String arrays[] = { "id MEDIUMINT UNSIGNED PRIMARY KEY AUTO_INCREMENT", "city VARCHAR(255)",
						"county VARCHAR(255)", "worksheetNum VARCHAR(255)", "worksheetTheme VARCHAR(255)",
						"worksheetType VARCHAR(255)", "operateType VARCHAR(255)", "currentLink VARCHAR(255)",
						"currentLinkHandler VARCHAR(255)", "sendBegTime VARCHAR(255)", "endTime VARCHAR(255)",
						"originator VARCHAR(255)", "installAddr VARCHAR(255)", "installAddrType VARCHAR(255)",
						"CRMorderNumber VARCHAR(255)", "acceptanceBusinessHall VARCHAR(255)",
						"receivingSalesPerson VARCHAR(255)", "acceptSalesmanPhone VARCHAR(255)",
						"customerName VARCHAR(255)", "customerContact VARCHAR(255)",
						"customerContactPhone VARCHAR(255)", "broadbandAccount VARCHAR(255)", "coverType VARCHAR(255)",
						"orderAcceptStartingTime VARCHAR(255)", "orderAcceptEndingTime VARCHAR(255)",
						"orderAcceptBacklogMan VARCHAR(255)", "orderAcceptHandler VARCHAR(255)",
						"openingBookingClosingTime VARCHAR(255)", "openingBookingBacklogMan VARCHAR(255)",
						"openingBookingHandler VARCHAR(255)", "onsiteOpeningStartingTime VARCHAR(255)",
						"onsiteOpeningBacklogMan VARCHAR(255)", "onsiteOpeningHandler VARCHAR(255)",
						"satisfactSurveyStartingTime VARCHAR(255)", "satisfactSurveyEndingTime VARCHAR(255)",
						"satisfactSurveyBacklogMan VARCHAR(255)", "satisfactSurveyHandler VARCHAR(255)",
						"subordinateBranch VARCHAR(255)", "daiweiCompany VARCHAR(255)", "maintenanceTeam VARCHAR(255)",
						"realConstructerUnit VARCHAR(255)", "realConstructer VARCHAR(255)",
						"activateAutomatically VARCHAR(255)", "duration VARCHAR(255)", "submitByPhone VARCHAR(255)",
						"villageName VARCHAR(255)", "STBLicensingParty VARCHAR(255)", "STPId VARCHAR(255)",
						"sendModel VARCHAR(255)", "replacementPreAddress VARCHAR(255)",
						"replacementAddress VARCHAR(255)", "lightCatSN VARCHAR(255)", "replacementPONID VARCHAR(255)",
						"replacementPrePONID VARCHAR(255)", "fileInTime VARCHAR(255)", "appointInTime VARCHAR(255)",
						"smallMicroBroadband VARCHAR(255)", "appointmentTime VARCHAR(255)", "daiWeiUnit VARCHAR(255)",
						"note VARCHAR(255)", "temporaryConstructer VARCHAR(255)",
						"temporaryConstructerPhone VARCHAR(255)", "orderReachingNetworkTime VARCHAR(255)",
						"outConstructionsFilingTime VARCHAR(255)", "villageId INT(11)" };
				OperateMysql.createTable(tabName, arrays);
				OperateMysql.truncateTable(tabName);
				firstCreate = false;
			}

			// 构造report
			TotalReport report = new TotalReport(str[j], str[j + 1], str[j + 2], str[j + 3], str[j + 4], str[j + 5],
					str[j + 6], str[j + 7], str[j + 8], str[j + 9], str[j + 10], str[j + 11], str[j + 12], str[j + 13],
					str[j + 14], str[j + 15], str[j + 16], str[j + 17], str[j + 18], str[j + 19], str[j + 20],
					str[j + 21], str[j + 22], str[j + 23], str[j + 24], str[j + 25], str[j + 26], str[j + 27],
					str[j + 28], str[j + 29], str[j + 30], str[j + 31], str[j + 32], str[j + 33], str[j + 34],
					str[j + 35], str[j + 36], str[j + 37], str[j + 38], str[j + 39], str[j + 40], str[j + 41],
					str[j + 42], str[j + 43], str[j + 44], str[j + 45], str[j + 46], str[j + 47], str[j + 48],
					str[j + 49], str[j + 50], str[j + 51], str[j + 52], str[j + 53], str[j + 54], str[j + 55],
					str[j + 56], str[j + 57], str[j + 58], str[j + 59], str[j + 60], str[j + 61], str[j + 62],
					villageId);

			// 向表中插入数据
			String[] names = { "city", "county", "worksheetNum", "worksheetTheme", "worksheetType", "operateType",
					"currentLink", "currentLinkHandler", "sendBegTime", "endTime", "originator", "installAddr",
					"installAddrType", "CRMorderNumber", "acceptanceBusinessHall", "receivingSalesPerson",
					"acceptSalesmanPhone", "customerName", "customerContact", "customerContactPhone",
					"broadbandAccount", "coverType", "orderAcceptStartingTime", "orderAcceptEndingTime",
					"orderAcceptBacklogMan", "orderAcceptHandler", "openingBookingClosingTime",
					"openingBookingBacklogMan", "openingBookingHandler", "onsiteOpeningStartingTime",
					"onsiteOpeningBacklogMan", "onsiteOpeningHandler", "satisfactSurveyStartingTime",
					"satisfactSurveyEndingTime", "satisfactSurveyBacklogMan", "satisfactSurveyHandler",
					"subordinateBranch", "daiweiCompany", "maintenanceTeam", "realConstructerUnit", "realConstructer",
					"activateAutomatically", "duration", "submitByPhone", "villageName", "STBLicensingParty", "STPId",
					"sendModel", "replacementPreAddress", "replacementAddress", "lightCatSN", "replacementPONID",
					"replacementPrePONID", "fileInTime", "appointInTime", "smallMicroBroadband", "appointmentTime",
					"daiWeiUnit", "note", "temporaryConstructer", "temporaryConstructerPhone",
					"orderReachingNetworkTime", "outConstructionsFilingTime", "villageId" };

			Object[] values = { report.getCity(), report.getCounty(), report.getWorksheetNum(),
					report.getWorksheetTheme(), report.getWorksheetType(), report.getOperateType(),
					report.getCurrentLink(), report.getCurrentLinkHandler(), report.getSendBegTime(),
					report.getEndTime(), report.getOriginator(), report.getInstallAddr(), report.getInstallAddrType(),
					report.getCRMorderNumber(), report.getAcceptanceBusinessHall(), report.getReceivingSalesPerson(),
					report.getAcceptSalesmanPhone(), report.getCustomerName(), report.getCustomerContact(),
					report.getCustomerContactPhone(), report.getBroadbandAccount(), report.getCoverType(),
					report.getOrderAcceptStartingTime(), report.getOrderAcceptEndingTime(),
					report.getOrderAcceptBacklogMan(), report.getOrderAcceptHandler(),
					report.getOpeningBookingClosingTime(), report.getOpeningBookingHandler(),
					report.getOpeningBookingHandler(), report.getOnsiteOpeningStartingTime(),
					report.getOnsiteOpeningBacklogMan(), report.getOnsiteOpeningHandler(),
					report.getSatisfactSurveyStartingTime(), report.getSatisfactSurveyEndingTime(),
					report.getSatisfactSurveyHandler(), report.getSatisfactSurveyBacklogMan(),
					report.getSubordinateBranch(), report.getDaiweiCompany(), report.getMaintenanceTeam(),
					report.getRealConstructerUnit(), report.getRealConstructer(), report.getActivateAutomatically(),
					report.getDuration(), report.getSubmitByPhone(), report.getVillageName(),
					report.getSTBLicensingParty(), report.getSTPId(), report.getSendModel(),
					report.getReplacementPreAddress(), report.getReplacementAddress(), report.getLightCatSN(),
					report.getReplacementPONID(), report.getReplacementPrePONID(), report.getFileInTime(),
					report.getAppointInTime(), report.getSmallMicroBroadband(), report.getAppointmentTime(),
					report.getDaiWeiUnit(), report.getNote(), report.getTemporaryConstructer(),
					report.getTemporaryConstructerPhone(), report.getOrderReachingNetworkTime(),
					report.getOutConstructionsFilingTime(), report.getVillageId() };
			
			OperateMysql.insertSignal(tabName, names, values);
		}

	}

	// 该方法专门用来打印excel提取的数据
	public static void printExecel(int curRow, List<String> cellList) {
		StringBuffer oneLineSb = new StringBuffer();
		oneLineSb.append("row" + curRow);
		oneLineSb.append("::");
		int curCol = 0;
		for (String cell : cellList) {
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