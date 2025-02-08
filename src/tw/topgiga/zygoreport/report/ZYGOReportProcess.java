package tw.topgiga.zygoreport.report;

import java.io.ByteArrayOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

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

		// 添加所有過濾條件
		if (p_DateFrom != null) {
			whereClause.add("m.updated >= " + DB.TO_DATE(p_DateFrom));
		}
		if (p_DateTo != null) {
			whereClause.add("m.updated <= " + DB.TO_DATE(p_DateTo));
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

		sql.append("WITH RankedData AS (").append(
				"  SELECT m.devicename AS devicename, m.groupname AS groupname, m.samplename AS samplename, m.operator AS operator, ")
				.append("         m.positionname AS positionname, md.dataname AS dataname, md.datavalue AS datavalue, ")
				.append("         ma.attributename AS attributename, ma.attributevalue AS attributevalue, ")
				.append("         md.updated AS updated, ").append("         ROW_NUMBER() OVER (PARTITION BY ")
				.append("             m.devicename, m.groupname, m.samplename, m.operator, ")
				.append("             m.positionname, md.dataname, md.datavalue, ")
				.append("             ma.attributename, ma.attributevalue ")
				.append("         ORDER BY m.updated DESC) as rn ").append("  FROM measure m ")
				.append("  LEFT JOIN measureddata md ON m.measure_id = md.measure_id ")
				.append("  LEFT JOIN measureattribute ma ON md.measureddata_id = ma.measureddata_id ")
				.append("  WHERE ").append(whereStr).append(") ")
				.append("SELECT devicename, groupname, samplename, operator, positionname, dataname, datavalue, attributename, attributevalue, updated ")
				.append("FROM RankedData WHERE rn = 1 ").append("ORDER BY updated DESC");

		return sql.toString();
	}

	private void generateReport(XSSFWorkbook workbook) throws Exception {
		XSSFSheet sheet = workbook.createSheet("ZYGO測量數據");

		// 建立樣式
		CellStyle headerStyle = createHeaderStyle(workbook);
		CellStyle basicStyle = createBasicStyle(workbook);
		CellStyle numberStyle = createNumberStyle(workbook);

		// 創建表頭
		Row headerRow = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(headers[i]);
			cell.setCellStyle(headerStyle);
		}

		PreparedStatement pstmt = null;
		ResultSet rs = null;
		int currentRow = 1;
		int totalRows = 0;

		try {
			// 首先獲取總行數
			String countSql = "SELECT COUNT(*) FROM (" + createSQL() + ") as total";
			pstmt = DB.prepareStatement(countSql, get_TrxName());
			rs = pstmt.executeQuery();
			if (rs.next()) {
				totalRows = rs.getInt(1);
			}
			DB.close(rs, pstmt);

			String batchSql = createSQL() + " OFFSET ? LIMIT ?";

			for (int offset = 0; offset < totalRows; offset += BATCH_SIZE) {
				pstmt = DB.prepareStatement(batchSql, get_TrxName());
				pstmt.setInt(1, offset);
				pstmt.setInt(2, BATCH_SIZE);
				rs = pstmt.executeQuery();

				List<String[]> batchData = new ArrayList<>();

				// 讀取這一批的數據
				while (rs.next()) {
					String[] rowData = new String[10]; // 改為10，因為現在有10個欄位
					rowData[0] = rs.getString("devicename") != null ? rs.getString("devicename") : "";
					rowData[1] = rs.getString("groupname") != null ? rs.getString("groupname") : "";
					rowData[2] = rs.getString("samplename") != null ? rs.getString("samplename") : "";
					rowData[3] = rs.getString("operator") != null ? rs.getString("operator") : "";
					rowData[4] = rs.getString("positionname") != null ? rs.getString("positionname") : "";
					rowData[5] = rs.getString("dataname") != null ? rs.getString("dataname") : "";
					rowData[6] = rs.getBigDecimal("datavalue") != null ? rs.getBigDecimal("datavalue").toString() : "0";
					rowData[7] = rs.getString("attributename") != null ? rs.getString("attributename") : "";
					rowData[8] = rs.getString("attributevalue") != null ? rs.getString("attributevalue") : "";
					rowData[9] = rs.getString("updated") != null ? rs.getString("updated") : ""; // 修正這行
					batchData.add(rowData);
				}
				DB.close(rs, pstmt);

				// 處理這一批數據
				for (int i = 0; i < batchData.size(); i++) {
					Row row = sheet.createRow(currentRow + i);
					String[] currentRowData = batchData.get(i);

					// 填充每個儲存格
					for (int j = 0; j < currentRowData.length; j++) {
						Cell cell = row.createCell(j);
						// 處理數值欄位
						if (j == 6 && currentRowData[j] != null) {
							try {
								cell.setCellValue(Double.parseDouble(currentRowData[j]));
								cell.setCellStyle(numberStyle);
							} catch (NumberFormatException e) {
								cell.setCellValue(currentRowData[j]);
								cell.setCellStyle(basicStyle);
							}
						} else {
							cell.setCellValue(currentRowData[j]);
							cell.setCellStyle(basicStyle);
						}
					}
				}

				// 合併相同值的儲存格
				for (int j = 0; j < headers.length; j++) {
					int startRow = currentRow;
					int endRow = currentRow;

					for (int i = 1; i < batchData.size(); i++) {
						String currentValue = batchData.get(i - 1)[j];
						String nextValue = batchData.get(i)[j];

						if (currentValue.equals(nextValue)) {
							endRow = currentRow + i;
						} else {
							if (startRow < endRow) {
								sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, j, j));
							}
							startRow = currentRow + i;
							endRow = currentRow + i;
						}
					}

					// 合併最後一段
					if (startRow < endRow) {
						sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, j, j));
					}
				}

				currentRow += batchData.size();
				batchData.clear();
			}

			// 自動調整欄寬
			for (int i = 0; i < headers.length; i++) {
				sheet.autoSizeColumn(i);
			}

			// 設置凍結窗格
			sheet.createFreezePane(0, 1);

		} finally {
			DB.close(rs, pstmt);
		}
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