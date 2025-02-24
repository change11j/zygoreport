package tw.topgiga.zygoreport.report;

import java.io.ByteArrayOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.compiere.model.MAttachment;
import org.compiere.model.MProcess;
import org.compiere.process.ProcessInfoParameter;
import org.compiere.process.SvrProcess;
import org.compiere.util.DB;

public class ZYGOReportProcess extends SvrProcess {
	private Timestamp p_DateFrom = null;
	private Timestamp p_DateTo = null;
	private String p_GroupName = null;
	private String p_operator = null;
	private String p_SampleName = null;
	private static final int BATCH_SIZE = 1000; // 每批處理的數據量
	// 定義表頭，增加操作者欄位
	String[] headers = { "儀器", "組別", "SampleName", "操作者", "點位", "量測欄位", "數值", "屬性", "屬性值", "時間" };

	@Override
	protected void prepare() {
		ProcessInfoParameter[] para = getParameter();
		for (int i = 0; i < para.length; i++) {
			String name = para[i].getParameterName();
			if (para[i].getParameter() == null)
				continue;

			if (name.equals("DateFrom"))
				p_DateFrom = (Timestamp) para[i].getParameter();
			else if (name.equals("DateTo"))
				p_DateTo = (Timestamp) para[i].getParameter();
			else if (name.equals("GroupName"))
				p_GroupName = para[i].getParameterAsString();
			else if (name.equals("operator"))
				p_operator = para[i].getParameterAsString();
			else if (name.equals("SampleName"))
				p_SampleName = para[i].getParameterAsString();
		}
	}

	@Override
	protected String doIt() throws Exception {
		try {
			XSSFWorkbook workbook = new XSSFWorkbook();
			generateReport(workbook);

			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			workbook.write(baos);
			workbook.close();

			int AD_Table_ID = getTable_ID();
			int Record_ID = getRecord_ID();

			if (AD_Table_ID > 0 && Record_ID > 0) {
				MAttachment attachment = MAttachment.get(getCtx(), AD_Table_ID, Record_ID);
				if (attachment == null)
					attachment = new MAttachment(getCtx(), AD_Table_ID, Record_ID, get_TrxName());

				attachment.addEntry("ZYGO測量報告.xlsx", baos.toByteArray());
				attachment.saveEx();

				return "報表已生成，請從附件下載";
			} else {
				MProcess process = new MProcess(getCtx(), getProcessInfo().getAD_Process_ID(), get_TrxName());
				MAttachment attachment = process.createAttachment();
				attachment.addEntry("ZYGO測量報告.xlsx", baos.toByteArray());
				attachment.saveEx();

				return "報表已生成，請從流程附件下載";
			}
		} catch (Exception e) {
			log.severe("Error generating report: " + e.getMessage());
			throw e;
		}
	}

	private String createSQL() {
		StringBuilder sql = new StringBuilder();
		ArrayList<String> whereClause = new ArrayList<>();

		// 基本條件
		whereClause.add("m.isactive = 'Y'");

		// 添加過濾條件
		if (p_DateFrom != null) {
			whereClause.add("md.updated >= " + DB.TO_DATE(p_DateFrom));
		}
		if (p_DateTo != null) {
			whereClause.add("md.updated <= " + DB.TO_DATE(p_DateTo));
		}
		if (p_GroupName != null && !p_GroupName.trim().isEmpty()) {
			whereClause.add("m.groupname = " + DB.TO_STRING(p_GroupName));
		}
		if (p_operator != null && !p_operator.trim().isEmpty()) {
			whereClause.add("m.operator = " + DB.TO_STRING(p_operator));
		}
		if (p_SampleName != null && !p_SampleName.trim().isEmpty()) {
			whereClause.add("m.samplename = " + DB.TO_STRING(p_SampleName));
		}

		String whereStr = String.join(" AND ", whereClause);

		sql.append("SELECT DISTINCT ").append("m.measure_id, ").append("COALESCE(m.slideid, '') as slideid, ")
				.append("COALESCE(m.devicename, '') as devicename, ").append("COALESCE(m.groupname, '') as groupname, ")
				.append("COALESCE(m.samplename, '') as samplename, ").append("COALESCE(m.operator, '') as operator, ")
				.append("COALESCE(m.positionname, '') as positionname, ")
				.append("COALESCE(md.dataname, '') as dataname, ").append("md.datavalue, ")
				.append("COALESCE(ma.attributename, '') as attributename, ")
				.append("COALESCE(ma.attributevalue, '') as attributevalue, ").append("m.updated ")
				.append("FROM measure m ").append("INNER JOIN measureddata md ON m.measure_id = md.measure_id ")
				.append("LEFT JOIN measureattribute ma ON md.measureddata_id = ma.measureddata_id ").append("WHERE ")
				.append(whereStr).append(" ORDER BY m.measure_id");

		return sql.toString();
	}

	private void generateReport(XSSFWorkbook workbook) throws Exception {
		XSSFSheet sheet = workbook.createSheet("ZYGO測量數據");

		// 建立樣式
		CellStyle headerStyle = createHeaderStyle(workbook);
		CellStyle basicStyle = createBasicStyle(workbook);
		CellStyle numberStyle = createNumberStyle(workbook);
		CellStyle groupHeaderStyle = createGroupHeaderStyle(workbook);

		Map<String, SlideData> slideDataMap = new TreeMap<>(Comparator.nullsLast(String::compareTo));

		// 收集數據部分
		try (PreparedStatement pstmt = DB.prepareStatement(createSQL(), get_TrxName());
				ResultSet rs = pstmt.executeQuery()) {

			String currentMeasureId = null;
			Map<String, String> currentAttributes = new TreeMap<>();

			while (rs.next()) {
				String slideId = rs.getString("slideid") != null ? rs.getString("slideid") : "";
				String measureId = rs.getString("measure_id");
				Timestamp updated = rs.getTimestamp("updated");

				if (currentMeasureId == null || !currentMeasureId.equals(measureId)) {
					currentAttributes.clear();
					currentMeasureId = measureId;
				}

				SlideData slideData = slideDataMap.computeIfAbsent(slideId, k -> {
					SlideData data = new SlideData();
					try {
						data.operator = rs.getString("operator") != null ? rs.getString("operator") : "";
						data.groupName = rs.getString("groupname") != null ? rs.getString("groupname") : "";
						data.sampleName = rs.getString("samplename") != null ? rs.getString("samplename") : "";
					} catch (SQLException e) {
						e.printStackTrace();
					}
					return data;
				});

				String attrName = rs.getString("attributename");
				String attrValue = rs.getString("attributevalue");
				if (attrName != null && !attrName.trim().isEmpty()) {
					currentAttributes.put(attrName, attrValue);
				}

				String positionName = rs.getString("positionname");
				String dataName = rs.getString("dataname");
				String dataValue = rs.getString("datavalue");

				if (positionName != null && !positionName.trim().isEmpty() && dataName != null
						&& !dataName.trim().isEmpty()) {
					slideData.addMeasurement(positionName, measureId, dataName, dataValue,
							new TreeMap<>(currentAttributes));
				}
			}
		}

		// 生成報表
		int rowNum = 0;

		// 對每個 slide 生成區塊
		for (Map.Entry<String, SlideData> slideEntry : slideDataMap.entrySet()) {
			String slideId = slideEntry.getKey();
			SlideData data = slideEntry.getValue();

			// 對每個屬性組合生成獨立的數據區塊
			for (Map.Entry<String, SlideData.GroupData> groupEntry : data.attributeGroups.entrySet()) {
				SlideData.GroupData groupData = groupEntry.getValue();

				int mergeColumns = Math.max(groupData.dataNames.size(), 2); // 最少合併到第2列

				// Slide ID 標題列
				Row slideRow = sheet.createRow(rowNum++);
				Cell slideCell = slideRow.createCell(0);
				slideCell.setCellValue("Slide_id");
				slideCell.setCellStyle(headerStyle);
				Cell slideValueCell = slideRow.createCell(1);
				slideValueCell.setCellValue(slideId);
				slideValueCell.setCellStyle(groupHeaderStyle);
				sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 1, mergeColumns));

				// 基本信息
				createInfoRow(sheet, rowNum++, "操作者", data.operator, basicStyle, headerStyle, mergeColumns);
				createInfoRow(sheet, rowNum++, "groupName", data.groupName, basicStyle, headerStyle, mergeColumns);

				// 屬性信息
				for (Map.Entry<String, String> attr : groupData.attributes.entrySet()) {
					createInfoRow(sheet, rowNum++, attr.getKey(), attr.getValue(), basicStyle, headerStyle,
							mergeColumns);
				}

				// 空白行
				rowNum++;

				// 生成數據表格
				rowNum = generateDataTable(sheet, rowNum, groupData, data.sampleName, headerStyle, basicStyle,
						numberStyle, mergeColumns);

				// 組之間添加兩個空白行
				rowNum += 2;
			}
		}

		// 自動調整列寬
		for (int i = 0; i < 7; i++) {
			sheet.autoSizeColumn(i);
		}

		// 設置凍結窗格
		sheet.createFreezePane(1, 1);
	}

	// 修改 generateDataTable 方法
	private int generateDataTable(XSSFSheet sheet, int startRow, SlideData.GroupData groupData, String sampleName,
			CellStyle headerStyle, CellStyle basicStyle, CellStyle numberStyle, int mergeColumns) {
		int rowNum = startRow;

		// SampleName 行
		Row sampleRow = sheet.createRow(rowNum++);
		Cell sampleCell = sampleRow.createCell(0);
		sampleCell.setCellValue("SampleName");
		sampleCell.setCellStyle(headerStyle);
		Cell sampleValueCell = sampleRow.createCell(1);
		sampleValueCell.setCellValue(sampleName);
		sampleValueCell.setCellStyle(basicStyle);
		sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 1, mergeColumns));

		// DataName 標題行
		Row dataNameRow = sheet.createRow(rowNum++);
		int colNum = 1;
		for (String dataName : groupData.dataNames) {
			Cell cell = dataNameRow.createCell(colNum++);
			cell.setCellValue(dataName);
			cell.setCellStyle(headerStyle);
		}

		// 數據行
		for (String position : groupData.measurements.keySet()) {
			for (Map.Entry<String, Map<String, String>> measureEntry : groupData.measurements.get(position)
					.entrySet()) {
				Row row = sheet.createRow(rowNum++);

				// 位置編號
				Cell posCell = row.createCell(0);
				posCell.setCellValue(position);
				posCell.setCellStyle(headerStyle);

				// 數據值
				colNum = 1;
				for (String dataName : groupData.dataNames) {
					Cell cell = row.createCell(colNum++);
					String value = measureEntry.getValue().get(dataName);
					if (value != null && !value.isEmpty()) {
						try {
							double numValue = Double.parseDouble(value);
							cell.setCellValue(numValue);
							cell.setCellStyle(numberStyle);
						} catch (NumberFormatException e) {
							cell.setCellValue(value);
							cell.setCellStyle(basicStyle);
						}
					}
				}
			}
		}

		return rowNum;
	}

//修改 createInfoRow 方法以支持合併儲存格
	// 創建信息行的方法確保參數正確
	private void createInfoRow(XSSFSheet sheet, int rowNum, String label, String value, CellStyle basicStyle,
			CellStyle headerStyle, int mergeColumns) {
		Row row = sheet.createRow(rowNum);
		Cell labelCell = row.createCell(0);
		labelCell.setCellValue(label);
		labelCell.setCellStyle(headerStyle);

		Cell valueCell = row.createCell(1);
		valueCell.setCellValue(value);
		valueCell.setCellStyle(basicStyle);

		sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 1, mergeColumns));
	}

// 輔助類別用於組織數據
	private static class SlideData {
		String operator = "";
		String groupName = "";
		String sampleName = "";

		// 用於存儲不同屬性組合的數據
		static class GroupData {
			Map<String, String> attributes = new TreeMap<>();
			Map<String, Map<String, Map<String, String>>> measurements = new TreeMap<>((a, b) -> {
				try {
					int numA = Integer.parseInt(a);
					int numB = Integer.parseInt(b);
					return Integer.compare(numA, numB);
				} catch (NumberFormatException e) {
					return a.compareTo(b);
				}
			});
			Set<String> dataNames = new TreeSet<>();
		}

		// 用於存儲不同的屬性組合組
		Map<String, GroupData> attributeGroups = new TreeMap<>();

		// 生成屬性組合的唯一標識符
		private String generateGroupKey(Map<String, String> attributes, String operator, String groupName) {
			StringBuilder key = new StringBuilder();
			key.append(operator).append("_").append(groupName);
			for (Map.Entry<String, String> entry : attributes.entrySet()) {
				key.append("_").append(entry.getKey()).append("=").append(entry.getValue());
			}
			return key.toString();
		}

		// 添加測量數據
		void addMeasurement(String position, String measureId, String dataName, String value,
				Map<String, String> currentAttributes) {
			String groupKey = generateGroupKey(currentAttributes, operator, groupName);

			GroupData groupData = attributeGroups.computeIfAbsent(groupKey, k -> {
				GroupData newGroup = new GroupData();
				newGroup.attributes.putAll(currentAttributes);
				return newGroup;
			});

			groupData.measurements.computeIfAbsent(position, k -> new TreeMap<>())
					.computeIfAbsent(measureId, k -> new TreeMap<>()).put(dataName, value != null ? value : "");
			groupData.dataNames.add(dataName);
		}
	}

// 群組標題樣式
	private CellStyle createGroupHeaderStyle(XSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font = workbook.createFont();
		font.setFontName("微軟正黑體");
		font.setBold(true);
		style.setFont(font);

		return style;
	}

	private CellStyle createHeaderStyle(XSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font = workbook.createFont();
		font.setFontName("微軟正黑體");
		font.setBold(true);
		style.setFont(font);

		return style;
	}

	private CellStyle createBasicStyle(XSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font = workbook.createFont();
		font.setFontName("微軟正黑體");
		style.setFont(font);

		return style;
	}

	private CellStyle createNumberStyle(XSSFWorkbook workbook) {
		CellStyle style = createBasicStyle(workbook);
		style.setDataFormat(workbook.createDataFormat().getFormat("#,##0.000"));
		return style;
	}

	private void createCell(Row row, int column, Object value, CellStyle style) {
		Cell cell = row.createCell(column);

		try {
			if (value == null) {
				cell.setCellValue("");
			} else if (value instanceof String) {
				cell.setCellValue((String) value);
			} else if (value instanceof java.math.BigDecimal) {
				double numValue = ((java.math.BigDecimal) value).doubleValue();
				cell.setCellValue(numValue);
			} else if (value instanceof Number) {
				cell.setCellValue(((Number) value).doubleValue());
			}
		} catch (Exception e) {
			log.warning("Error setting cell value at column " + column + ": " + e.getMessage());
			cell.setCellValue("");
		}
		cell.setCellStyle(style);
	}
}