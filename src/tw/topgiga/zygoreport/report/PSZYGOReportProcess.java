package tw.topgiga.zygoreport.report;

import java.io.ByteArrayOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
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

public class PSZYGOReportProcess extends SvrProcess {
	private Timestamp p_DateFrom = null;
	private Timestamp p_DateTo = null;
	private String p_GroupName = null;
	private String p_operator = null;
	private String p_SampleName = null;
	private String p_SlideId = null;

	// 存儲slideId與sampleName的映射
	private Map<String, String> sampleNameMap = new HashMap<>();

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
			else if (name.equals("SlideId"))
				p_SlideId = para[i].getParameterAsString();
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

				attachment.addEntry("ZYGO_PS報表.xlsx", baos.toByteArray());
				attachment.saveEx();

				return "PS報表已生成，請從附件下載";
			} else {
				MProcess process = new MProcess(getCtx(), getProcessInfo().getAD_Process_ID(), get_TrxName());
				MAttachment attachment = process.createAttachment();
				attachment.addEntry("ZYGO_PS報表.xlsx", baos.toByteArray());
				attachment.saveEx();

				return "PS報表已生成，請從流程附件下載";
			}
		} catch (Exception e) {
			log.severe("生成報表時發生錯誤: " + e.getMessage());
			throw e;
		}
	}

	// SQL查詢 - 獲取屬性數據（排除HT、DOM和Dose）
	private String createAttributesSQL() {
		StringBuilder sql = new StringBuilder();
		ArrayList<String> whereClause = new ArrayList<>();

		// 基本條件
		whereClause.add("m.isactive = 'Y'");

		// 添加過濾條件
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
		if (p_SlideId != null && !p_SlideId.trim().isEmpty()) {
			whereClause.add("m.slideid = " + DB.TO_STRING(p_SlideId));
		}

		String whereStr = String.join(" AND ", whereClause);

		sql.append("SELECT ").append("m.measure_id, ").append("COALESCE(m.slideid, '') as slideid, ")
				.append("COALESCE(m.samplename, '') as samplename, ") // 額外獲取SampleName
				.append("COALESCE(ma.attributename, '') as attributename, ")
				.append("COALESCE(ma.attributevalue, '') as attributevalue ").append("FROM measure m ")
				.append("LEFT JOIN measureddata md ON m.measure_id = md.measure_id ")
				.append("LEFT JOIN measureattribute ma ON md.measureddata_id = ma.measureddata_id ").append("WHERE ")
				.append(whereStr).append(" AND UPPER(ma.attributename) NOT IN ('HT', 'DOM', 'DOSE') ")
				.append("ORDER BY m.slideid, m.measure_id");

		return sql.toString();
	}

	// SQL查詢 - 獲取Dose、HT和DOM屬性
	private String createDoseHTDomSQL() {
		StringBuilder sql = new StringBuilder();
		ArrayList<String> whereClause = new ArrayList<>();

		// 基本條件
		whereClause.add("m.isactive = 'Y'");

		// 添加過濾條件
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
		if (p_SlideId != null && !p_SlideId.trim().isEmpty()) {
			whereClause.add("m.slideid = " + DB.TO_STRING(p_SlideId));
		}

		// 添加屬性名稱條件到主要 WHERE 子句
		whereClause.add("(UPPER(ma.attributename) IN ('HT', 'DOM', 'DOSE') OR ma.attributename = 'dose')");

		String whereStr = String.join(" AND ", whereClause);

		sql.append("SELECT ").append("m.measure_id, ").append("COALESCE(m.slideid, '') as slideid, ")
				.append("COALESCE(m.samplename, '') as samplename, ") // 額外獲取SampleName
				.append("UPPER(ma.attributename) as attributename, ")
				.append("COALESCE(ma.attributevalue, '') as attributevalue ").append("FROM measure m ")
				.append("LEFT JOIN measureddata md ON m.measure_id = md.measure_id ")
				.append("LEFT JOIN measureattribute ma ON md.measureddata_id = ma.measureddata_id ").append("WHERE ")
				.append(whereStr).append(" ORDER BY m.slideid, m.measure_id");

		return sql.toString();
	}

	// SQL查詢 - 獲取測量數據
	private String createMeasurementDataSQL() {
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
		if (p_SlideId != null && !p_SlideId.trim().isEmpty()) {
			whereClause.add("m.slideid = " + DB.TO_STRING(p_SlideId));
		}

		String whereStr = String.join(" AND ", whereClause);

		sql.append("SELECT ").append("m.measure_id, ").append("COALESCE(m.slideid, '') as slideid, ")
				.append("COALESCE(m.positionname, '') as positionname, ")
				.append("COALESCE(md.dataname, '') as dataname, ").append("md.datavalue ").append("FROM measure m ")
				.append("INNER JOIN measureddata md ON m.measure_id = md.measure_id ").append("WHERE ").append(whereStr)
				.append(" AND md.dataname IN ('TCD DX-95%', 'TCD DY', 'PS-BOT-DX', 'PS-BOT-DY', 'PS-Hp') ")
				.append("ORDER BY m.slideid, m.positionname, m.measure_id");

		return sql.toString();
	}

	// 生成報表
	private void generateReport(XSSFWorkbook workbook) throws Exception {
		XSSFSheet sheet = workbook.createSheet("ZYGO_PS測量數據");

		// 建立樣式
		CellStyle headerStyle = createHeaderStyle(workbook);
		CellStyle basicStyle = createBasicStyle(workbook);
		CellStyle numberStyle = createNumberStyle(workbook, "#,##0.00"); // 改為2位小數
		CellStyle blueStyle = createBlueStyle(workbook);
		CellStyle orangeStyle = createOrangeStyle(workbook);

		// 1. 獲取屬性數據 (不包含HT、DOM和Dose)
		// 結構: slideId -> measureId -> attributeName -> attributeValue
		Map<String, Map<String, Map<String, String>>> attributesMap = new HashMap<>();
		try (PreparedStatement pstmt = DB.prepareStatement(createAttributesSQL(), get_TrxName());
				ResultSet rs = pstmt.executeQuery()) {

			while (rs.next()) {
				String slideId = rs.getString("slideid") != null ? rs.getString("slideid") : "";
				String sampleName = rs.getString("samplename") != null ? rs.getString("samplename") : "";
				String measureId = rs.getString("measure_id");
				String attrName = rs.getString("attributename");
				String attrValue = rs.getString("attributevalue") != null ? rs.getString("attributevalue") : "";

				// 保存slideId和sampleName的映射
				sampleNameMap.put(slideId, sampleName);

				if (attrName != null && !attrName.trim().isEmpty()) {
					Map<String, Map<String, String>> measureMap = attributesMap.computeIfAbsent(slideId,
							k -> new HashMap<>());
					Map<String, String> attrMap = measureMap.computeIfAbsent(measureId, k -> new HashMap<>());
					attrMap.put(attrName, attrValue);
				}
			}
		}

		// 2. 獲取HT、DOM和Dose數據
		// 結構: slideId -> measureId -> attributeName -> attributeValue
		Map<String, Map<String, Map<String, String>>> htDomDoseMap = new HashMap<>();
		try (PreparedStatement pstmt = DB.prepareStatement(createDoseHTDomSQL(), get_TrxName());
				ResultSet rs = pstmt.executeQuery()) {

			while (rs.next()) {
				String slideId = rs.getString("slideid") != null ? rs.getString("slideid") : "";
				String sampleName = rs.getString("samplename") != null ? rs.getString("samplename") : "";
				String measureId = rs.getString("measure_id");
				String attrName = rs.getString("attributename");
				String attrValue = rs.getString("attributevalue") != null ? rs.getString("attributevalue") : "";

				// 保存slideId和sampleName的映射
				sampleNameMap.put(slideId, sampleName);

				if (attrName != null && !attrName.trim().isEmpty()) {
					Map<String, Map<String, String>> measureMap = htDomDoseMap.computeIfAbsent(slideId,
							k -> new HashMap<>());
					Map<String, String> attrMap = measureMap.computeIfAbsent(measureId, k -> new HashMap<>());
					attrMap.put(attrName, attrValue);
				}
			}
		}

		// 3. 獲取測量數據
		// 結構: slideId -> HT-DOM-Key -> positionName -> dataName -> value
		Map<String, Map<String, Map<String, Map<String, String>>>> reorganizedData = new HashMap<>();
		Map<String, List<String>> slideHtDomKeysMap = new HashMap<>();

		try (PreparedStatement pstmt = DB.prepareStatement(createMeasurementDataSQL(), get_TrxName());
				ResultSet rs = pstmt.executeQuery()) {

			// 收集所有數據并存儲到臨時結構中
			Map<String, Map<String, Map<String, Map<String, String>>>> rawData = new HashMap<>();

			while (rs.next()) {
				String slideId = rs.getString("slideid") != null ? rs.getString("slideid") : "";
				String measureId = rs.getString("measure_id");
				String positionName = rs.getString("positionname") != null ? rs.getString("positionname") : "";
				String dataName = rs.getString("dataname");
				String dataValue = rs.getString("datavalue") != null ? rs.getString("datavalue") : "";

				// 存儲原始數據
				Map<String, Map<String, Map<String, String>>> measureMap = rawData.computeIfAbsent(slideId,
						k -> new HashMap<>());
				Map<String, Map<String, String>> positionMap = measureMap.computeIfAbsent(measureId,
						k -> new HashMap<>());
				Map<String, String> dataMap = positionMap.computeIfAbsent(positionName, k -> new HashMap<>());
				dataMap.put(dataName, dataValue);
				dataMap.put("measure_id", measureId); // 保存measure_id便於後續參考
			}

			log.info("原始數據收集完成，重組為HT-DOM結構");

			// 為每個slideId處理HT-DOM分組
			for (String slideId : rawData.keySet()) {
				Map<String, Map<String, Map<String, String>>> measureMap = rawData.get(slideId);

				// 為這個slideId創建一個measureId到HT-DOM鍵的映射
				Map<String, String> measureToHtDomMap = new HashMap<>();

				// 首先，獲取每個measureId的HT-DOM值並分組
				for (String measureId : measureMap.keySet()) {
					Map<String, Map<String, String>> htDomMapForSlide = htDomDoseMap.getOrDefault(slideId,
							new HashMap<>());
					Map<String, String> htDomValues = htDomMapForSlide.getOrDefault(measureId, new HashMap<>());

					// 獲取HT值並轉換為整數
					String htRaw = htDomValues.getOrDefault("HT", "100");
					String ht;
					try {
						double htDouble = Double.parseDouble(htRaw);
						ht = String.valueOf((int) htDouble);
					} catch (NumberFormatException e) {
						ht = htRaw;
					}

					// 修改為:
					String domRaw = htDomValues.getOrDefault("DOM", "");
					String dom;
					try {
						double domDouble = Double.parseDouble(domRaw);
						dom = String.valueOf((int) domDouble);
					} catch (NumberFormatException e) {
						dom = domRaw;
					}
					String htDomKey = ht + "-" + dom;

					measureToHtDomMap.put(measureId, htDomKey);

					// 記錄這個slideId的所有HT-DOM鍵
					List<String> htDomKeys = slideHtDomKeysMap.computeIfAbsent(slideId, k -> new ArrayList<>());
					if (!htDomKeys.contains(htDomKey)) {
						htDomKeys.add(htDomKey);
					}
				}

				// 對HT-DOM鍵進行排序 (HT降序，DOM升序)
				List<String> htDomKeys = slideHtDomKeysMap.get(slideId);
				Collections.sort(htDomKeys, (a, b) -> {
					String[] partsA = a.split("-");
					String[] partsB = b.split("-");

					double htA = 0, htB = 0;
					try {
						htA = Double.parseDouble(partsA[0]);
						htB = Double.parseDouble(partsB[0]);
					} catch (NumberFormatException e) {
						log.warning("無法解析HT值: " + e.getMessage());
					}

					// 先按HT降序排列
					int htCompare = Double.compare(htB, htA);
					if (htCompare != 0)
						return htCompare;

					// 如果HT相同，按DOM升序排列
					String domA = partsA.length > 1 ? partsA[1] : "";
					String domB = partsB.length > 1 ? partsB[1] : "";
					return domA.compareTo(domB);
				});

				log.info("SlideID: " + slideId + " 的HT-DOM鍵排序後: " + htDomKeys);

				// 根據排序後的HT-DOM鍵重組數據
				Map<String, Map<String, Map<String, String>>> htDomMapData = reorganizedData.computeIfAbsent(slideId,
						k -> new LinkedHashMap<>());

				// 預先為所有HT-DOM鍵創建空Map，確保順序與排序一致
				for (String htDomKey : htDomKeys) {
					htDomMapData.put(htDomKey, new HashMap<>());
				}

				// 填充HT-DOM數據結構
				for (String measureId : measureMap.keySet()) {
					String htDomKey = measureToHtDomMap.get(measureId);
					Map<String, Map<String, String>> positionMap = measureMap.get(measureId);

					// 獲取此HT-DOM鍵對應的位置Map
					Map<String, Map<String, String>> htDomPositionMap = htDomMapData.get(htDomKey);

					for (String positionName : positionMap.keySet()) {
						Map<String, String> dataMap = positionMap.get(positionName);
						Map<String, String> existingDataMap = htDomPositionMap.get(positionName);

						if (existingDataMap == null) {
							// 如果還沒有此position的數據，直接添加
							htDomPositionMap.put(positionName, new HashMap<>(dataMap));
						} else {
							// 如果已有數據，檢查是否要更新
							String existingMeasureId = existingDataMap.getOrDefault("measure_id", "-1");
							if (Integer.parseInt(measureId) > Integer.parseInt(existingMeasureId)) {
								htDomPositionMap.put(positionName, new HashMap<>(dataMap));
								log.info("更新SlideID: " + slideId + ", HT-DOM: " + htDomKey + ", Position: "
										+ positionName + ", 從MeasureID " + existingMeasureId + " 到 " + measureId);
							}
						}
					}
				}
			}

			// 檢查reorganizedData是否包含所有數據
			for (String slideId : reorganizedData.keySet()) {
				Map<String, Map<String, Map<String, String>>> htDomMapData = reorganizedData.get(slideId);
				log.info("SlideID: " + slideId + " 的HT-DOM數據統計:");

				for (String htDomKey : htDomMapData.keySet()) {
					Map<String, Map<String, String>> positionMap = htDomMapData.get(htDomKey);
					log.info("  HT-DOM鍵: " + htDomKey + ", 包含 " + positionMap.size() + " 個position");

					for (String position : positionMap.keySet()) {
						Map<String, String> dataMap = positionMap.get(position);
						log.info("    Position " + position + ": 包含 " + dataMap.size() + " 個數據項，MeasureID="
								+ dataMap.getOrDefault("measure_id", "未知"));
					}
				}
			}
		}

		// 4. 生成報表 - 改為水平排列而不是垂直疊加
		int maxColumn = 0;
		int currentCol = 0;

		// 取得所有slideId並排序，保證報表順序一致
		List<String> slideIds = new ArrayList<>(reorganizedData.keySet());
		Collections.sort(slideIds);

		for (String slideId : slideIds) {
			Map<String, Map<String, Map<String, String>>> htDomData = reorganizedData.get(slideId);
			List<String> htDomKeys = slideHtDomKeysMap.getOrDefault(slideId, new ArrayList<>());
			Map<String, Map<String, String>> attributesByMeasure = attributesMap.getOrDefault(slideId, new HashMap<>());
			Map<String, Map<String, String>> htDomDoseByMeasure = htDomDoseMap.getOrDefault(slideId, new HashMap<>());

			// 計算此SlideID需要的列數
			int columnsNeeded = htDomKeys.size() * 2 + 1; // 每個HT-DOM組佔2列(X和Y)

			// 如果此slideId需要的列數會超出工作表寬度，則換行
			if (currentCol + columnsNeeded > 256) { // Excel列數上限
				currentCol = 0;
			}

			// 為此SlideID創建報表
			int columnsUsed = createSlideReportNew(workbook, sheet, 0, currentCol, slideId, htDomKeys, htDomData,
					htDomDoseByMeasure, attributesByMeasure, headerStyle, basicStyle, numberStyle, orangeStyle,
					blueStyle);

			// 更新當前列位置和最大列位置
			currentCol += columnsUsed; // +1 表示添加一列間隔
			maxColumn = Math.max(maxColumn, currentCol);
		}

		// 設定合適的列寬
		for (int i = 0; i < maxColumn; i++) {
			sheet.setColumnWidth(i, 12 * 200); // 設置適當的列寬
		}
		// 計算 DOM 行的位置
		int freezeRow = 0;
		for (Row row : sheet) {
			Cell firstCell = row.getCell(0);
			if (firstCell != null && firstCell.getCellType() == CellType.STRING) {
				String cellValue = firstCell.getStringCellValue();
				if ("Dom".equals(cellValue)) {
					// 找到 DOM 行，保存其索引 (+1 是因為我們要凍結 DOM 行之後的行)
					freezeRow = row.getRowNum() + 2; // +2 表示包括 X/Y 行
					break;
				}
			}
		}

		// 如果找到了 DOM 行，設置凍結窗格
		if (freezeRow > 0) {
			sheet.createFreezePane(0, freezeRow);
		}
	}

	// 為單個SlideID生成完整報表（水平布局）
	private int createSlideReportNew(XSSFWorkbook workbook, XSSFSheet sheet, int rowNum, int startCol, String slideId,
			List<String> htDomKeys, Map<String, Map<String, Map<String, String>>> htDomData,
			Map<String, Map<String, String>> htDomDoseByMeasure, Map<String, Map<String, String>> attributesByMeasure,
			CellStyle headerStyle, CellStyle basicStyle, CellStyle numberStyle, CellStyle orangeStyle,
			CellStyle blueStyle) {

		int columnCount = htDomKeys.size() * 2 + 1; // 每個HT-DOM組佔2列(X和Y)
		int currentCol = startCol;

		// 1. 創建Item行 - 使用sampleName代替slideId
		String sampleName = sampleNameMap.getOrDefault(slideId, slideId);
		Row itemRow = sheet.getRow(rowNum);
		if (itemRow == null) {
			itemRow = sheet.createRow(rowNum);
		}
		rowNum++;

		Cell itemLabelCell = itemRow.createCell(currentCol);
		itemLabelCell.setCellValue("Item");
		itemLabelCell.setCellStyle(headerStyle);

		Cell itemValueCell = itemRow.createCell(currentCol + 1);
		itemValueCell.setCellValue(sampleName);
		itemValueCell.setCellStyle(basicStyle);

		// 跨列合併Item值
		if (htDomKeys.size() > 1) {
			sheet.addMergedRegion(
					new CellRangeAddress(rowNum - 1, rowNum - 1, currentCol + 1, currentCol + columnCount - 1));
			// 設置每個格子的樣式
			for (int i = 2; i < columnCount; i++) {
				Cell cell = itemRow.createCell(currentCol + i);
				cell.setCellStyle(basicStyle);
			}
		}

		// 2. 創建動態屬性行
		if (!htDomKeys.isEmpty() && !htDomData.isEmpty()) {
			// 從第一個HT-DOM鍵的第一個position獲取一個measureId
			String firstHtDomKey = htDomKeys.get(0);
			Map<String, Map<String, String>> positionData = htDomData.get(firstHtDomKey);
			if (positionData != null && !positionData.isEmpty()) {
				String firstPosition = positionData.keySet().iterator().next();
				String measureId = positionData.get(firstPosition).getOrDefault("measure_id", "");

				if (!measureId.isEmpty()) {
					// 獲取該measureId的所有屬性
					Map<String, String> attrs = attributesByMeasure.getOrDefault(measureId, new HashMap<>());

					// 為每個屬性創建一行（排除Mask的藍色樣式）
					for (String attrName : attrs.keySet()) {
						Row attrRow = sheet.getRow(rowNum);
						if (attrRow == null) {
							attrRow = sheet.createRow(rowNum);
						}
						rowNum++;

						Cell attrLabelCell = attrRow.createCell(currentCol);
						attrLabelCell.setCellValue(attrName);
						attrLabelCell.setCellStyle(headerStyle);

						// 檢查所有HT-DOM組中該屬性的值是否相同
						boolean allSameValue = true;
						String firstAttrValue = null;

						for (int i = 0; i < htDomKeys.size(); i++) {
							String htDomKey = htDomKeys.get(i);
							Map<String, Map<String, String>> htDomPositionData = htDomData.get(htDomKey);

							if (!htDomPositionData.isEmpty()) {
								String posKey = htDomPositionData.keySet().iterator().next();
								String mid = htDomPositionData.get(posKey).getOrDefault("measure_id", "");

								if (!mid.isEmpty()) {
									Map<String, String> attrMap = attributesByMeasure.getOrDefault(mid,
											new HashMap<>());
									String attrValue = attrMap.getOrDefault(attrName, "");

									if (firstAttrValue == null) {
										firstAttrValue = attrValue;
									} else if (!firstAttrValue.equals(attrValue)) {
										allSameValue = false;
										break;
									}
								}
							}
						}

						// 如果所有值相同，則合併單元格
						if (allSameValue && firstAttrValue != null && htDomKeys.size() > 1) {
							Cell valueCell = attrRow.createCell(currentCol + 1);
							valueCell.setCellValue(firstAttrValue);
							valueCell.setCellStyle(basicStyle);

							// 合併橫跨所有HT-DOM組的單元格
							sheet.addMergedRegion(new CellRangeAddress(attrRow.getRowNum(), attrRow.getRowNum(),
									currentCol + 1, currentCol + htDomKeys.size() * 2));

							// 設置所有單元格的樣式
							for (int i = 2; i <= htDomKeys.size() * 2; i++) {
								Cell cell = attrRow.createCell(currentCol + i);
								cell.setCellStyle(basicStyle);
							}
						} else {
							// 為每個HT-DOM組填充屬性值
							for (int i = 0; i < htDomKeys.size(); i++) {
								String htDomKey = htDomKeys.get(i);
								positionData = htDomData.get(htDomKey);

								if (!positionData.isEmpty()) {
									// 獲取第一個position的measureId
									String posKey = positionData.keySet().iterator().next();
									String mid = positionData.get(posKey).getOrDefault("measure_id", "");

									if (!mid.isEmpty()) {
										// 獲取該measureId的屬性值
										attrs = attributesByMeasure.getOrDefault(mid, new HashMap<>());
										String attrValue = attrs.getOrDefault(attrName, "");

										// 使用基本樣式（不再使用Mask藍色樣式）
										CellStyle valueStyle = basicStyle;

										// 填充值（每個HT-DOM組佔用2列）
										Cell attrValueCell = attrRow.createCell(currentCol + i * 2 + 1);
										attrValueCell.setCellValue(attrValue);
										attrValueCell.setCellStyle(valueStyle);

										// 跨兩列合併單元格
										sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1,
												currentCol + i * 2 + 1, currentCol + i * 2 + 2));

										// 設置第二列單元格樣式
										Cell mergedCell = attrRow.createCell(currentCol + i * 2 + 2);
										mergedCell.setCellStyle(valueStyle);
									}
								}
							}
						}
					}
				}
			}
		}

		// 3. 創建Dose行 - 使用橘底紅字樣式
		Row doseRow = sheet.getRow(rowNum);
		if (doseRow == null) {
			doseRow = sheet.createRow(rowNum);
		}
		rowNum++;

		Cell doseLabelCell = doseRow.createCell(currentCol);
		doseLabelCell.setCellValue("Dose (mJ)");
		doseLabelCell.setCellStyle(headerStyle);

		// 獲取Dose值
		String doseValue = "";
		if (!htDomKeys.isEmpty()) {
			String firstHtDomKey = htDomKeys.get(0);
			Map<String, Map<String, String>> positionData = htDomData.get(firstHtDomKey);
			if (positionData != null && !positionData.isEmpty()) {
				String firstPosition = positionData.keySet().iterator().next();
				String measureId = positionData.get(firstPosition).getOrDefault("measure_id", "");

				if (!measureId.isEmpty()) {
					Map<String, String> htDomDose = htDomDoseByMeasure.getOrDefault(measureId, new HashMap<>());
					doseValue = htDomDose.getOrDefault("DOSE", "");
					// 檢查小寫dose
					if (doseValue.isEmpty()) {
						doseValue = htDomDose.getOrDefault("dose", "");
					}
				}
			}
		}

		// 創建Dose樣式
		CellStyle doseStyle = workbook.createCellStyle();
		doseStyle.cloneStyleFrom(basicStyle);
		doseStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		doseStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		Font doseFont = workbook.createFont();
		doseFont.setFontName("微軟正黑體");
		doseFont.setColor(IndexedColors.RED.getIndex());
		doseStyle.setFont(doseFont);

		// 檢查所有HT-DOM組中Dose的值是否相同
		boolean allSameDose = true;
		String firstDoseValue = doseValue;

		for (int i = 1; i < htDomKeys.size(); i++) {
			String htDomKey = htDomKeys.get(i);
			Map<String, Map<String, String>> positionData = htDomData.get(htDomKey);

			if (!positionData.isEmpty()) {
				String posKey = positionData.keySet().iterator().next();
				String mid = positionData.get(posKey).getOrDefault("measure_id", "");

				if (!mid.isEmpty()) {
					Map<String, String> htDomDose = htDomDoseByMeasure.getOrDefault(mid, new HashMap<>());
					String dose = htDomDose.getOrDefault("DOSE", "");
					if (dose.isEmpty()) {
						dose = htDomDose.getOrDefault("dose", "");
					}

					if (!firstDoseValue.equals(dose)) {
						allSameDose = false;
						break;
					}
				}
			}
		}

		// 如果所有Dose值相同，則合併單元格
		if (allSameDose && !firstDoseValue.isEmpty() && htDomKeys.size() > 1) {
			Cell doseValueCell = doseRow.createCell(currentCol + 1);
			doseValueCell.setCellValue(firstDoseValue);
			doseValueCell.setCellStyle(doseStyle);

			// 合併橫跨所有HT-DOM組的單元格
			sheet.addMergedRegion(new CellRangeAddress(doseRow.getRowNum(), doseRow.getRowNum(), currentCol + 1,
					currentCol + htDomKeys.size() * 2));

			// 設置所有單元格的樣式
			for (int i = 2; i <= htDomKeys.size() * 2; i++) {
				Cell cell = doseRow.createCell(currentCol + i);
				cell.setCellStyle(doseStyle);
			}
		} else {
			// 為每個HT-DOM組填充Dose值
			for (int i = 0; i < htDomKeys.size(); i++) {
				String htDomKey = htDomKeys.get(i);
				Map<String, Map<String, String>> positionData = htDomData.get(htDomKey);

				if (!positionData.isEmpty()) {
					String posKey = positionData.keySet().iterator().next();
					String mid = positionData.get(posKey).getOrDefault("measure_id", "");

					if (!mid.isEmpty()) {
						Map<String, String> htDomDose = htDomDoseByMeasure.getOrDefault(mid, new HashMap<>());
						String dose = htDomDose.getOrDefault("DOSE", "");
						if (dose.isEmpty()) {
							dose = htDomDose.getOrDefault("dose", "");
						}

						// 填充值（每個HT-DOM組佔用2列）
						Cell doseValueCell = doseRow.createCell(currentCol + i * 2 + 1);
						doseValueCell.setCellValue(dose);
						doseValueCell.setCellStyle(doseStyle);

						// 跨兩列合併單元格
						sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, currentCol + i * 2 + 1,
								currentCol + i * 2 + 2));

						// 設置第二列單元格樣式
						Cell mergedCell = doseRow.createCell(currentCol + i * 2 + 2);
						mergedCell.setCellStyle(doseStyle);
					}
				}
			}
		}

		// 4. 創建HT (%)行
		Row htRow = sheet.getRow(rowNum);
		if (htRow == null) {
			htRow = sheet.createRow(rowNum);
		}
		rowNum++;

		Cell htLabelCell = htRow.createCell(currentCol);
		htLabelCell.setCellValue("HT (%)");
		htLabelCell.setCellStyle(headerStyle);

		// 5. 創建Dom行
		Row domRow = sheet.getRow(rowNum);
		if (domRow == null) {
			domRow = sheet.createRow(rowNum);
		}
		rowNum++;

		Cell domLabelCell = domRow.createCell(currentCol);
		domLabelCell.setCellValue("Dom");
		domLabelCell.setCellStyle(headerStyle);

		log.info("填充SlideID: " + slideId + " 的HT-DOM數據，共 " + htDomKeys.size() + " 組");

		// 計算相同HT值的DOM數量
		Map<String, List<Integer>> htIndices = new HashMap<>();
		for (int i = 0; i < htDomKeys.size(); i++) {
			String htDomKey = htDomKeys.get(i);
			String[] parts = htDomKey.split("-");
			String ht = parts[0];

			if (!htIndices.containsKey(ht)) {
				htIndices.put(ht, new ArrayList<>());
			}
			htIndices.get(ht).add(i);
		}

		// 填充HT和Dom數據，並按相同HT值合併
		for (String ht : htIndices.keySet()) {
			List<Integer> indices = htIndices.get(ht);
			int startIndex = indices.get(0);
			int endIndex = indices.get(indices.size() - 1);

			// HT值 - 合併相同HT的所有列
			int startHtCol = currentCol + startIndex * 2 + 1;
			int endHtCol = currentCol + endIndex * 2 + 2;

			Cell htValueCell = htRow.createCell(startHtCol);
			htValueCell.setCellValue(ht);
			htValueCell.setCellStyle(orangeStyle);

			if (startHtCol < endHtCol) {
				sheet.addMergedRegion(new CellRangeAddress(htRow.getRowNum(), htRow.getRowNum(), startHtCol, endHtCol));

				// 設置合併區域內的所有單元格樣式
				for (int col = startHtCol + 1; col <= endHtCol; col++) {
					Cell mergedCell = htRow.createCell(col);
					mergedCell.setCellStyle(orangeStyle);
				}
			}

			// 針對每個DOM值單獨處理
			for (int idx : indices) {
				String htDomKey = htDomKeys.get(idx);
				String[] parts = htDomKey.split("-");
				String dom = parts.length > 1 ? parts[1] : "";

				// Dom值 - 每個DOM值佔2列(X和Y)
				int xCol = currentCol + idx * 2 + 1;
				int yCol = currentCol + idx * 2 + 2;

				Cell domValueCell = domRow.createCell(xCol);
				domValueCell.setCellValue(dom);
				domValueCell.setCellStyle(orangeStyle);

				sheet.addMergedRegion(new CellRangeAddress(domRow.getRowNum(), domRow.getRowNum(), xCol, yCol));

				Cell domMergedCell = domRow.createCell(yCol);
				domMergedCell.setCellStyle(orangeStyle);
			}
		}

		// 6. 創建X/Y行
		Row xyRow = sheet.getRow(rowNum);
		if (xyRow == null) {
			xyRow = sheet.createRow(rowNum);
		}
		rowNum++;

		Cell emptyCellXY = xyRow.createCell(currentCol);
		emptyCellXY.setCellStyle(headerStyle);

		for (int i = 0; i < htDomKeys.size(); i++) {
			Cell xCell = xyRow.createCell(currentCol + i * 2 + 1);
			xCell.setCellValue("X");
			xCell.setCellStyle(basicStyle);

			Cell yCell = xyRow.createCell(currentCol + i * 2 + 2);
			yCell.setCellValue("Y");
			yCell.setCellStyle(basicStyle);
		}

		// 7. 獲取所有position並排序
		Set<String> allPositions = new HashSet<>();
		for (String htDomKey : htDomKeys) {
			Map<String, Map<String, String>> positionData = htDomData.getOrDefault(htDomKey, new HashMap<>());
			allPositions.addAll(positionData.keySet());
		}

		List<String> positions = new ArrayList<>(allPositions);
		// 數字排序
		Collections.sort(positions, (a, b) -> {
			try {
				return Integer.compare(Integer.parseInt(a), Integer.parseInt(b));
			} catch (NumberFormatException e) {
				return a.compareTo(b);
			}
		});

		log.info("Position排序結果: " + String.join(", ", positions));

		// 8. 創建平均值行
		// TCD平均值行
		Row tcdAvgRow = sheet.getRow(rowNum);
		if (tcdAvgRow == null) {
			tcdAvgRow = sheet.createRow(rowNum);
		}
		rowNum++;

		Cell tcdAvgLabelCell = tcdAvgRow.createCell(currentCol);
		tcdAvgLabelCell.setCellValue("TCD");
		tcdAvgLabelCell.setCellStyle(headerStyle);

		// BCD平均值行
		Row bcdAvgRow = sheet.getRow(rowNum);
		if (bcdAvgRow == null) {
			bcdAvgRow = sheet.createRow(rowNum);
		}
		rowNum++;

		Cell bcdAvgLabelCell = bcdAvgRow.createCell(currentCol);
		bcdAvgLabelCell.setCellValue("BCD");
		bcdAvgLabelCell.setCellStyle(headerStyle);

		// PSH平均值行
		Row pshAvgRow = sheet.getRow(rowNum);
		if (pshAvgRow == null) {
			pshAvgRow = sheet.createRow(rowNum);
		}
		rowNum++;

		Cell pshAvgLabelCell = pshAvgRow.createCell(currentCol);
		pshAvgLabelCell.setCellValue("PSH");
		pshAvgLabelCell.setCellStyle(headerStyle);

		// B-T行
		Row btRow = sheet.getRow(rowNum);
		if (btRow == null) {
			btRow = sheet.createRow(rowNum);
		}
		rowNum++;

		Cell btLabelCell = btRow.createCell(currentCol);
		btLabelCell.setCellValue("B-T (Bot-Top)");
		btLabelCell.setCellStyle(blueStyle);

		// 識別具有HT=100的DOM值，用於創建M-S行
		List<String> mainDomValues = new ArrayList<>();
		for (String htDomKey : htDomKeys) {
			String[] parts = htDomKey.split("-");
			String ht = parts[0];
			String dom = parts.length > 1 ? parts[1] : "";

			if (ht.equals("100")) {
				mainDomValues.add(dom);
			}
		}

		// 創建M-S行
		Map<String, Integer> mainDomRowMap = new HashMap<>(); // 存儲每個main dom對應的行號
		for (String mainDom : mainDomValues) {
			Row msRow = sheet.getRow(rowNum);
			if (msRow == null) {
				msRow = sheet.createRow(rowNum);
			}

			Cell msLabelCell = msRow.createCell(currentCol);
			msLabelCell.setCellValue("M(" + mainDom + ")-S");
			msLabelCell.setCellStyle(blueStyle);
			mainDomRowMap.put(mainDom, rowNum);
			rowNum++;
		}

		// 為所有HT-DOM組計算平均值並填充到相應單元格
		Map<String, Double[]> htDomAverages = new HashMap<>(); // 存儲每個HT-DOM組的平均值 [tcdX, tcdY, bcdX, bcdY, psh]

		for (int i = 0; i < htDomKeys.size(); i++) {
			String htDomKey = htDomKeys.get(i);
			Map<String, Map<String, String>> positionDataMap = htDomData.getOrDefault(htDomKey, new HashMap<>());

			int xCol = currentCol + i * 2 + 1;
			int yCol = currentCol + i * 2 + 2;

			// 計算該HT-DOM組下所有position的平均值
			double tcdXSum = 0, tcdYSum = 0, bcdXSum = 0, bcdYSum = 0, pshSum = 0;
			int tcdXCount = 0, tcdYCount = 0, bcdXCount = 0, bcdYCount = 0, pshCount = 0;

			for (String position : positions) {
				Map<String, String> dataMap = positionDataMap.getOrDefault(position, new HashMap<>());

				if (!dataMap.isEmpty()) {
					// 獲取數據
					String tcdX = dataMap.getOrDefault("TCD DX-95%", "");
					String tcdY = dataMap.getOrDefault("TCD DY", "");
					String bcdX = dataMap.getOrDefault("PS-BOT-DX", "");
					String bcdY = dataMap.getOrDefault("PS-BOT-DY", "");
					String psh = dataMap.getOrDefault("PS-Hp", "");

					// 累加有效數據
					if (!tcdX.isEmpty()) {
						try {
							tcdXSum += Double.parseDouble(tcdX);
							tcdXCount++;
						} catch (NumberFormatException e) {
							log.warning("無法解析TCD DX值: " + tcdX);
						}
					}

					if (!tcdY.isEmpty()) {
						try {
							tcdYSum += Double.parseDouble(tcdY);
							tcdYCount++;
						} catch (NumberFormatException e) {
							log.warning("無法解析TCD DY值: " + tcdY);
						}
					}

					if (!bcdX.isEmpty()) {
						try {
							bcdXSum += Double.parseDouble(bcdX);
							bcdXCount++;
						} catch (NumberFormatException e) {
							log.warning("無法解析PS-BOT-DX值: " + bcdX);
						}
					}

					if (!bcdY.isEmpty()) {
						try {
							bcdYSum += Double.parseDouble(bcdY);
							bcdYCount++;
						} catch (NumberFormatException e) {
							log.warning("無法解析PS-BOT-DY值: " + bcdY);
						}
					}

					if (!psh.isEmpty()) {
						try {
							pshSum += Double.parseDouble(psh);
							pshCount++;
						} catch (NumberFormatException e) {
							log.warning("無法解析PS-Hp值: " + psh);
						}
					}
				}
			}

			// 計算平均值
			double tcdXAvg = tcdXCount > 0 ? tcdXSum / tcdXCount : 0;
			double tcdYAvg = tcdYCount > 0 ? tcdYSum / tcdYCount : 0;
			double bcdXAvg = bcdXCount > 0 ? bcdXSum / bcdXCount : 0;
			double bcdYAvg = bcdYCount > 0 ? bcdYSum / bcdYCount : 0;
			double pshAvg = pshCount > 0 ? pshSum / pshCount : 0;

			// 存儲平均值以供後續計算
			htDomAverages.put(htDomKey, new Double[] { tcdXAvg, tcdYAvg, bcdXAvg, bcdYAvg, pshAvg });

			// 填充平均值到表格
			// TCD平均值
			Cell tcdXAvgCell = tcdAvgRow.createCell(xCol);
			tcdXAvgCell.setCellValue(tcdXAvg);
			tcdXAvgCell.setCellStyle(numberStyle);

			Cell tcdYAvgCell = tcdAvgRow.createCell(yCol);
			tcdYAvgCell.setCellValue(tcdYAvg);
			tcdYAvgCell.setCellStyle(numberStyle);

			// BCD平均值
			Cell bcdXAvgCell = bcdAvgRow.createCell(xCol);
			bcdXAvgCell.setCellValue(bcdXAvg);
			bcdXAvgCell.setCellStyle(numberStyle);

			Cell bcdYAvgCell = bcdAvgRow.createCell(yCol);
			bcdYAvgCell.setCellValue(bcdYAvg);
			bcdYAvgCell.setCellStyle(numberStyle);

			// PSH平均值 - 合併X和Y列
			Cell pshAvgCell = pshAvgRow.createCell(xCol);
			pshAvgCell.setCellValue(pshAvg);
			pshAvgCell.setCellStyle(numberStyle);

			// 合併PSH的X和Y列
			sheet.addMergedRegion(new CellRangeAddress(pshAvgRow.getRowNum(), pshAvgRow.getRowNum(), xCol, yCol));

			// 確保Y列有樣式
			Cell pshAvgYCell = pshAvgRow.createCell(yCol);
			pshAvgYCell.setCellStyle(numberStyle);

			// 計算B-T (BCD-TCD)的差值
			double btXDiff = bcdXAvg - tcdXAvg;
			double btYDiff = bcdYAvg - tcdYAvg;

			// 填充B-T差值
			Cell btXCell = btRow.createCell(xCol);
			btXCell.setCellValue(btXDiff);
			btXCell.setCellStyle(numberStyle);

			Cell btYCell = btRow.createCell(yCol);
			btYCell.setCellValue(btYDiff);
			btYCell.setCellStyle(numberStyle);
		}

		// 填充M-S行
		for (int i = 0; i < htDomKeys.size(); i++) {
			String htDomKey = htDomKeys.get(i);
			String[] parts = htDomKey.split("-");
			String ht = parts[0];
			String dom = parts.length > 1 ? parts[1] : "";

			int xCol = currentCol + i * 2 + 1;
			int yCol = currentCol + i * 2 + 2;

			// 獲取當前HT-DOM組的PSH平均值
			Double[] currentAvgs = htDomAverages.get(htDomKey);
			double currentPsh = currentAvgs != null ? currentAvgs[4] : 0;

			// 對於每個main dom (HT=100的DOM值)，計算M-S差值
			for (String mainDom : mainDomValues) {
				int msRowIndex = mainDomRowMap.get(mainDom);
				Row msRow = sheet.getRow(msRowIndex);

				// 找到對應的main HT-DOM組
				String mainHtDomKey = "100-" + mainDom;
				Double[] mainAvgs = htDomAverages.get(mainHtDomKey);

				if (ht.equals("100") && dom.equals(mainDom)) {
					// 如果當前就是main dom，則顯示"-"
					Cell msXCell = msRow.createCell(xCol);
					msXCell.setCellValue("-");
					msXCell.setCellStyle(basicStyle);

					// 合併X和Y列
					sheet.addMergedRegion(new CellRangeAddress(msRowIndex, msRowIndex, xCol, yCol));

					Cell msYCell = msRow.createCell(yCol);
					msYCell.setCellStyle(basicStyle);
				} else {
					// 否則計算PSH差值 (main PSH - current PSH)
					// 否則計算PSH差值 (main PSH - current PSH)
					double mainPsh = mainAvgs != null ? mainAvgs[4] : 0;
					double pshDiff = mainPsh - currentPsh;

					// 在所有HT=100的DOM情況下，都不計算M-S差值
					if (ht.equals("100")) {
						Cell msXCell = msRow.createCell(xCol);
						msXCell.setCellValue("-");
						msXCell.setCellStyle(basicStyle);

						// 合併X和Y列
						sheet.addMergedRegion(new CellRangeAddress(msRowIndex, msRowIndex, xCol, yCol));

						Cell msYCell = msRow.createCell(yCol);
						msYCell.setCellStyle(basicStyle);
					} else {
						Cell msXCell = msRow.createCell(xCol);
						msXCell.setCellValue(pshDiff);
						msXCell.setCellStyle(numberStyle);

						// 合併X和Y列
						sheet.addMergedRegion(new CellRangeAddress(msRowIndex, msRowIndex, xCol, yCol));

						Cell msYCell = msRow.createCell(yCol);
						msYCell.setCellStyle(numberStyle);
					}
				}
			}
		}

		// 9. 處理每個Position的數據行
		rowNum = createPositionRows(sheet, rowNum, currentCol, slideId, htDomKeys, htDomData, positions, mainDomValues,
				htDomAverages, headerStyle, basicStyle, numberStyle, orangeStyle, blueStyle);

		return columnCount; // 返回此報表使用的列數
	}

	// 處理每個Position的資料行
	private int createPositionRows(XSSFSheet sheet, int rowNum, int startCol, String slideId, List<String> htDomKeys,
			Map<String, Map<String, Map<String, String>>> htDomData, List<String> positions, List<String> mainDomValues,
			Map<String, Double[]> htDomAverages, CellStyle headerStyle, CellStyle basicStyle, CellStyle numberStyle,
			CellStyle orangeStyle, CellStyle blueStyle) {

		// 為每個 Position 創建數據行
		for (String position : positions) {
			// 創建 Position 標題行
			Row positionRow = sheet.getRow(rowNum);
			if (positionRow == null) {
				positionRow = sheet.createRow(rowNum);
			}
			rowNum++;

			Cell positionCell = positionRow.createCell(startCol);
			positionCell.setCellValue("Position " + position);
			positionCell.setCellStyle(headerStyle);

			// 合併標題單元格
			int endCol = startCol + htDomKeys.size() * 2;
			for (int i = startCol + 1; i <= endCol; i++) {
				Cell cell = positionRow.createCell(i);
				cell.setCellStyle(headerStyle);
			}
			sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, startCol, endCol));

			// 創建 TCD 行
			Row tcdRow = sheet.getRow(rowNum);
			if (tcdRow == null) {
				tcdRow = sheet.createRow(rowNum);
			}
			rowNum++;

			Cell tcdLabelCell = tcdRow.createCell(startCol);
			tcdLabelCell.setCellValue("TCD");
			tcdLabelCell.setCellStyle(headerStyle);

			// 創建 BCD 行
			Row bcdRow = sheet.getRow(rowNum);
			if (bcdRow == null) {
				bcdRow = sheet.createRow(rowNum);
			}
			rowNum++;

			Cell bcdLabelCell = bcdRow.createCell(startCol);
			bcdLabelCell.setCellValue("BCD");
			bcdLabelCell.setCellStyle(headerStyle);

			// 創建 PSH 行
			Row pshRow = sheet.getRow(rowNum);
			if (pshRow == null) {
				pshRow = sheet.createRow(rowNum);
			}
			rowNum++;

			Cell pshLabelCell = pshRow.createCell(startCol);
			pshLabelCell.setCellValue("PSH");
			pshLabelCell.setCellStyle(headerStyle);

			// 為每個 main dom 創建 M-S 行
			Map<String, Integer> mainDomRowMap = new HashMap<>(); // 存儲每個 main dom 對應的行號
			for (String mainDom : mainDomValues) {
				Row msRow = sheet.getRow(rowNum);
				if (msRow == null) {
					msRow = sheet.createRow(rowNum);
				}

				Cell msLabelCell = msRow.createCell(startCol);
				msLabelCell.setCellValue("M(" + mainDom + ")-S");
				msLabelCell.setCellStyle(blueStyle);
				mainDomRowMap.put(mainDom, rowNum);
				rowNum++;
			}

			// 填充每個 HT-DOM 組的數據
			for (int i = 0; i < htDomKeys.size(); i++) {
				String htDomKey = htDomKeys.get(i);
				String[] parts = htDomKey.split("-");
				String ht = parts[0];
				String dom = parts.length > 1 ? parts[1] : "";

				Map<String, Map<String, String>> positionDataMap = htDomData.getOrDefault(htDomKey, new HashMap<>());
				Map<String, String> dataMap = positionDataMap.getOrDefault(position, new HashMap<>());

				int xCol = startCol + i * 2 + 1;
				int yCol = startCol + i * 2 + 2;

				// 填充 TCD, BCD, PSH 數據
				if (!dataMap.isEmpty()) {
					String tcdX = dataMap.getOrDefault("TCD DX-95%", "");
					String tcdY = dataMap.getOrDefault("TCD DY", "");
					String bcdX = dataMap.getOrDefault("PS-BOT-DX", "");
					String bcdY = dataMap.getOrDefault("PS-BOT-DY", "");
					String psh = dataMap.getOrDefault("PS-Hp", "");

					// TCD 數據
					setNumericCellValue(tcdRow.createCell(xCol), tcdX, numberStyle);
					setNumericCellValue(tcdRow.createCell(yCol), tcdY, numberStyle);

					// BCD 數據
					setNumericCellValue(bcdRow.createCell(xCol), bcdX, numberStyle);
					setNumericCellValue(bcdRow.createCell(yCol), bcdY, numberStyle);

					// PSH 數據 - 合併X和Y列
					Cell pshCell = pshRow.createCell(xCol);
					setNumericCellValue(pshCell, psh, numberStyle);
					sheet.addMergedRegion(new CellRangeAddress(pshRow.getRowNum(), pshRow.getRowNum(), xCol, yCol));

					// 確保Y列有樣式
					Cell pshYCell = pshRow.createCell(yCol);
					pshYCell.setCellStyle(numberStyle);

					// 獲取當前 position 的 PSH 值
					double currentPsh = 0;
					try {
						if (!psh.isEmpty()) {
							currentPsh = Double.parseDouble(psh);
						}
					} catch (NumberFormatException e) {
						log.warning("無法解析 PSH 值: " + psh);
					}

					// 填充 M-S 行
					for (String mainDom : mainDomValues) {
						int msRowIndex = mainDomRowMap.get(mainDom);
						Row msRow = sheet.getRow(msRowIndex);

						// 找到對應的 main HT-DOM 組的 position 數據
						String mainHtDomKey = "100-" + mainDom;
						Map<String, Map<String, String>> mainPositionDataMap = htDomData.getOrDefault(mainHtDomKey,
								new HashMap<>());
						Map<String, String> mainDataMap = mainPositionDataMap.getOrDefault(position, new HashMap<>());

						if (ht.equals("100") && dom.equals(mainDom)) {
							// 如果當前就是 main dom，則顯示"-"
							Cell msXCell = msRow.createCell(xCol);
							msXCell.setCellValue("-");
							msXCell.setCellStyle(basicStyle);

							// 合併X和Y列
							sheet.addMergedRegion(new CellRangeAddress(msRowIndex, msRowIndex, xCol, yCol));

							Cell msYCell = msRow.createCell(yCol);
							msYCell.setCellStyle(basicStyle);
						} else {
							// 如果HT=100，不論DOM值如何，都不計算差值
							if (ht.equals("100")) {
								Cell msXCell = msRow.createCell(xCol);
								msXCell.setCellValue("-");
								msXCell.setCellStyle(basicStyle);

								// 合併X和Y列
								sheet.addMergedRegion(new CellRangeAddress(msRowIndex, msRowIndex, xCol, yCol));

								Cell msYCell = msRow.createCell(yCol);
								msYCell.setCellStyle(basicStyle);
							} else {
								// 獲取 main position 的 PSH 值
								double mainPsh = 0;
								String mainPshStr = mainDataMap.getOrDefault("PS-Hp", "");
								try {
									if (!mainPshStr.isEmpty()) {
										mainPsh = Double.parseDouble(mainPshStr);
									}
								} catch (NumberFormatException e) {
									log.warning("無法解析 main PSH 值: " + mainPshStr);
								}

								// 計算 PSH 差值 (main PSH - current PSH)
								double pshDiff = mainPsh - currentPsh;

								Cell msXCell = msRow.createCell(xCol);
								msXCell.setCellValue(pshDiff);
								msXCell.setCellStyle(numberStyle);

								// 合併X和Y列
								sheet.addMergedRegion(new CellRangeAddress(msRowIndex, msRowIndex, xCol, yCol));

								Cell msYCell = msRow.createCell(yCol);
								msYCell.setCellStyle(numberStyle);
							}
						}
					}
				} else {
					// 如果沒有數據，填充空單元格
					tcdRow.createCell(xCol).setCellStyle(numberStyle);
					tcdRow.createCell(yCol).setCellStyle(numberStyle);
					bcdRow.createCell(xCol).setCellStyle(numberStyle);
					bcdRow.createCell(yCol).setCellStyle(numberStyle);

					// PSH行 - 合併X和Y列
					Cell pshCell = pshRow.createCell(xCol);
					pshCell.setCellStyle(numberStyle);
					sheet.addMergedRegion(new CellRangeAddress(pshRow.getRowNum(), pshRow.getRowNum(), xCol, yCol));

					Cell pshYCell = pshRow.createCell(yCol);
					pshYCell.setCellStyle(numberStyle);

					// 填充 M-S 行
					for (String mainDom : mainDomValues) {
						int msRowIndex = mainDomRowMap.get(mainDom);
						Row msRow = sheet.getRow(msRowIndex);

						if (ht.equals("100") && dom.equals(mainDom)) {
							Cell msXCell = msRow.createCell(xCol);
							msXCell.setCellValue("-");
							msXCell.setCellStyle(basicStyle);

							// 合併X和Y列
							sheet.addMergedRegion(new CellRangeAddress(msRowIndex, msRowIndex, xCol, yCol));

							Cell msYCell = msRow.createCell(yCol);
							msYCell.setCellStyle(basicStyle);
						} else if (ht.equals("100")) {
							// 如果HT=100，不管DOM值是什麼，都不計算差值
							Cell msXCell = msRow.createCell(xCol);
							msXCell.setCellValue("-");
							msXCell.setCellStyle(basicStyle);

							// 合併X和Y列
							sheet.addMergedRegion(new CellRangeAddress(msRowIndex, msRowIndex, xCol, yCol));

							Cell msYCell = msRow.createCell(yCol);
							msYCell.setCellStyle(basicStyle);
						} else {
							msRow.createCell(xCol).setCellStyle(numberStyle);

							// 合併X和Y列
							sheet.addMergedRegion(new CellRangeAddress(msRowIndex, msRowIndex, xCol, yCol));

							msRow.createCell(yCol).setCellStyle(numberStyle);
						}
					}
				}
			}
		}

		return rowNum;
	}

	// 輔助方法設置數值單元格
	private void setNumericCellValue(Cell cell, String value, CellStyle style) {
		if (value != null && !value.isEmpty()) {
			try {
				double numValue = Double.parseDouble(value);
				cell.setCellValue(numValue);
			} catch (NumberFormatException e) {
				cell.setCellValue(value);
			}
		}
		cell.setCellStyle(style);
	}

	// 用於存儲測量數據的輔助類
	private class MeasurementData {
		String measureId;
		String dataName;
		String dataValue;

		MeasurementData(String measureId, String dataName, String dataValue) {
			this.measureId = measureId;
			this.dataName = dataName;
			this.dataValue = dataValue;
		}
	}

	// 創建標題樣式
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

	// 創建基本樣式
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

	// 創建數字樣式
	private CellStyle createNumberStyle(XSSFWorkbook workbook, String format) {
		CellStyle style = createBasicStyle(workbook);
		style.setDataFormat(workbook.createDataFormat().getFormat(format));
		return style;
	}

	// 創建橙色背景樣式
	private CellStyle createOrangeStyle(XSSFWorkbook workbook) {
		CellStyle style = createBasicStyle(workbook);
		style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return style;
	}

	// 創建藍色文字樣式
	private CellStyle createBlueStyle(XSSFWorkbook workbook) {
		CellStyle style = createBasicStyle(workbook);

		Font font = workbook.createFont();
		font.setFontName("微軟正黑體");
		font.setColor(IndexedColors.BLUE.getIndex());
		style.setFont(font);

		return style;
	}

	private CellStyle createDoseStyle(XSSFWorkbook workbook) {
		CellStyle style = createBasicStyle(workbook);
		style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		Font font = workbook.createFont();
		font.setFontName("微軟正黑體");
		font.setColor(IndexedColors.RED.getIndex());
		style.setFont(font);

		return style;
	}
}