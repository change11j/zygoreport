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

	// SQL查詢 - 獲取屬性數據（排除HT和DOM）
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
				.append("COALESCE(ma.attributename, '') as attributename, ")
				.append("COALESCE(ma.attributevalue, '') as attributevalue ").append("FROM measure m ")
				.append("LEFT JOIN measureddata md ON m.measure_id = md.measure_id ")
				.append("LEFT JOIN measureattribute ma ON md.measureddata_id = ma.measureddata_id ").append("WHERE ")
				.append(whereStr).append(" AND UPPER(ma.attributename) NOT IN ('HT', 'DOM') ")
				.append("ORDER BY m.slideid, m.measure_id");

		return sql.toString();
	}

	// SQL查詢 - 獲取HT和DOM屬性
	private String createHTDomSQL() {
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
				.append("UPPER(ma.attributename) as attributename, ")
				.append("COALESCE(ma.attributevalue, '') as attributevalue ").append("FROM measure m ")
				.append("LEFT JOIN measureddata md ON m.measure_id = md.measure_id ")
				.append("LEFT JOIN measureattribute ma ON md.measureddata_id = ma.measureddata_id ").append("WHERE ")
				.append(whereStr).append(" AND UPPER(ma.attributename) IN ('HT', 'DOM') ")
				.append("ORDER BY m.slideid, m.measure_id");

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
		CellStyle numberStyle = createNumberStyle(workbook);
		CellStyle orangeStyle = createOrangeStyle(workbook);
		CellStyle blueStyle = createBlueStyle(workbook);

		// 1. 獲取屬性數據 (不包含HT和DOM)
		// 結構: slideId -> measureId -> attributeName -> attributeValue
		Map<String, Map<String, Map<String, String>>> attributesMap = new HashMap<>();
		try (PreparedStatement pstmt = DB.prepareStatement(createAttributesSQL(), get_TrxName());
				ResultSet rs = pstmt.executeQuery()) {

			while (rs.next()) {
				String slideId = rs.getString("slideid") != null ? rs.getString("slideid") : "";
				String measureId = rs.getString("measure_id");
				String attrName = rs.getString("attributename");
				String attrValue = rs.getString("attributevalue") != null ? rs.getString("attributevalue") : "";

				if (attrName != null && !attrName.trim().isEmpty()) {
					Map<String, Map<String, String>> measureMap = attributesMap.computeIfAbsent(slideId,
							k -> new HashMap<>());
					Map<String, String> attrMap = measureMap.computeIfAbsent(measureId, k -> new HashMap<>());
					attrMap.put(attrName, attrValue);
				}
			}
		}

		// 2. 獲取HT和DOM數據
		// 結構: slideId -> measureId -> attributeName -> attributeValue
		Map<String, Map<String, Map<String, String>>> htDomMap = new HashMap<>();
		try (PreparedStatement pstmt = DB.prepareStatement(createHTDomSQL(), get_TrxName());
				ResultSet rs = pstmt.executeQuery()) {

			while (rs.next()) {
				String slideId = rs.getString("slideid") != null ? rs.getString("slideid") : "";
				String measureId = rs.getString("measure_id");
				String attrName = rs.getString("attributename");
				String attrValue = rs.getString("attributevalue") != null ? rs.getString("attributevalue") : "";

				if (attrName != null && !attrName.trim().isEmpty()) {
					Map<String, Map<String, String>> measureMap = htDomMap.computeIfAbsent(slideId,
							k -> new HashMap<>());
					Map<String, String> attrMap = measureMap.computeIfAbsent(measureId, k -> new HashMap<>());
					attrMap.put(attrName, attrValue);
				}
			}
		}

		// 3. 獲取測量數據
		// 結構: slideId -> HT-DOM-Key -> positionName -> dataName -> value
		// 改進 1: 在數據讀取部分加入更多檢查和日誌
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
					Map<String, Map<String, String>> htDomMapForSlide = htDomMap.getOrDefault(slideId, new HashMap<>());
					Map<String, String> htDomValues = htDomMapForSlide.getOrDefault(measureId, new HashMap<>());
					String ht = htDomValues.getOrDefault("HT", "100.0");
					String dom = htDomValues.getOrDefault("DOM", "");
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

		// 4. 生成報表
		int rowNum = 0;

		for (String slideId : reorganizedData.keySet()) {
			Map<String, Map<String, Map<String, String>>> htDomData = reorganizedData.get(slideId);
			List<String> htDomKeys = slideHtDomKeysMap.getOrDefault(slideId, new ArrayList<>());
			Map<String, Map<String, String>> attributesByMeasure = attributesMap.getOrDefault(slideId, new HashMap<>());

			// 空白行分隔不同SlideID
			if (rowNum > 0) {
				rowNum += 2;
			}

			// 計算真正需要的列數
			int columnCount = htDomKeys.size() * 2 + 1; // 每個HT-DOM組佔2列(X和Y)

			// 為此SlideID創建報表
			rowNum = createSlideReportNew(sheet, rowNum, slideId, htDomKeys, htDomData, attributesByMeasure,
					headerStyle, basicStyle, numberStyle, orangeStyle, blueStyle);
		}
	}

	// 為單個SlideID生成完整報表（使用新的數據結構）
	private int createSlideReportNew(XSSFSheet sheet, int rowNum, String slideId, List<String> htDomKeys,
			Map<String, Map<String, Map<String, String>>> htDomData,
			Map<String, Map<String, String>> attributesByMeasure, CellStyle headerStyle, CellStyle basicStyle,
			CellStyle numberStyle, CellStyle orangeStyle, CellStyle blueStyle) {

		int startRow = rowNum;
		int columnCount = htDomKeys.size() * 2 + 1; // 每個HT-DOM組佔2列(X和Y)

		// 1. 創建Item行
		Row itemRow = sheet.createRow(rowNum++);
		Cell itemLabelCell = itemRow.createCell(0);
		itemLabelCell.setCellValue("Item");
		itemLabelCell.setCellStyle(headerStyle);

		Cell itemValueCell = itemRow.createCell(1);
		itemValueCell.setCellValue(slideId);
		itemValueCell.setCellStyle(basicStyle);

		// 跨列合併Item值
		if (htDomKeys.size() > 1) {
			sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 1, columnCount - 1));
			// 設置每個格子的樣式
			for (int i = 2; i < columnCount; i++) {
				Cell cell = itemRow.createCell(i);
				cell.setCellStyle(basicStyle);
			}
		}

		// 2. 創建動態屬性行
		if (!htDomKeys.isEmpty() && !htDomData.isEmpty()) {
			// 從第一個HT-DOM鍵的第一個position獲取一個measureId
			String firstHtDomKey = htDomKeys.get(0);
			Map<String, Map<String, String>> positionData = htDomData.get(firstHtDomKey);
			String firstPosition = positionData.keySet().iterator().next();
			String measureId = positionData.get(firstPosition).getOrDefault("measure_id", "");

			if (!measureId.isEmpty()) {
				// 獲取該measureId的所有屬性
				Map<String, String> attrs = attributesByMeasure.getOrDefault(measureId, new HashMap<>());

				// 為每個屬性創建一行
				for (String attrName : attrs.keySet()) {
					Row attrRow = sheet.createRow(rowNum++);
					Cell attrLabelCell = attrRow.createCell(0);
					attrLabelCell.setCellValue(attrName);
					attrLabelCell.setCellStyle(headerStyle);

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

								// 選擇樣式（Mask使用藍色）
								CellStyle valueStyle = attrName.equals("Mask") ? blueStyle : basicStyle;

								// 填充值（每個HT-DOM組佔用2列）
								Cell attrValueCell = attrRow.createCell(i * 2 + 1);
								attrValueCell.setCellValue(attrValue);
								attrValueCell.setCellStyle(valueStyle);

								// 跨兩列合併單元格
								sheet.addMergedRegion(
										new CellRangeAddress(rowNum - 1, rowNum - 1, i * 2 + 1, i * 2 + 2));

								// 設置第二列單元格樣式
								Cell mergedCell = attrRow.createCell(i * 2 + 2);
								mergedCell.setCellStyle(valueStyle);
							}
						}
					}
				}
			}
		}

		// 3. 創建HT (%)行
		Row htRow = sheet.createRow(rowNum++);
		Cell htLabelCell = htRow.createCell(0);
		htLabelCell.setCellValue("HT (%)");
		htLabelCell.setCellStyle(headerStyle);

		// 4. 創建Dom行
		Row domRow = sheet.createRow(rowNum++);
		Cell domLabelCell = domRow.createCell(0);
		domLabelCell.setCellValue("Dom");
		domLabelCell.setCellStyle(headerStyle);

		log.info("填充SlideID: " + slideId + " 的HT-DOM數據，共 " + htDomKeys.size() + " 組");

		// 填充HT和Dom數據
		for (int i = 0; i < htDomKeys.size(); i++) {
			String htDomKey = htDomKeys.get(i);
			String[] parts = htDomKey.split("-");
			String ht = parts[0];
			String dom = parts.length > 1 ? parts[1] : "";
			log.info("  處理HT-DOM鍵: " + htDomKey + ", HT=" + ht + ", DOM=" + dom);

			// HT值
			Cell htValueCell = htRow.createCell(i * 2 + 1);
			htValueCell.setCellValue(ht);
			htValueCell.setCellStyle(orangeStyle);

			sheet.addMergedRegion(new CellRangeAddress(rowNum - 2, rowNum - 2, i * 2 + 1, i * 2 + 2));
			Cell htMergedCell = htRow.createCell(i * 2 + 2);
			htMergedCell.setCellStyle(orangeStyle);

			// Dom值
			Cell domValueCell = domRow.createCell(i * 2 + 1);
			domValueCell.setCellValue(dom);
			domValueCell.setCellStyle(orangeStyle);

			sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, i * 2 + 1, i * 2 + 2));
			Cell domMergedCell = domRow.createCell(i * 2 + 2);
			domMergedCell.setCellStyle(orangeStyle);
		}

		// 5. 創建X/Y行
		Row xyRow = sheet.createRow(rowNum++);
		Cell emptyCellXY = xyRow.createCell(0);
		emptyCellXY.setCellStyle(headerStyle);

		for (int i = 0; i < htDomKeys.size(); i++) {
			Cell xCell = xyRow.createCell(i * 2 + 1);
			xCell.setCellValue("X");
			xCell.setCellStyle(basicStyle);

			Cell yCell = xyRow.createCell(i * 2 + 2);
			yCell.setCellValue("Y");
			yCell.setCellStyle(basicStyle);
		}

		// 6. 獲取所有position並排序
		Set<String> allPositions = new HashSet<>();
		for (String htDomKey : htDomKeys) {
			Map<String, Map<String, String>> positionData = htDomData.getOrDefault(htDomKey, new HashMap<>());
			allPositions.addAll(positionData.keySet());
		}

		List<String> positions = new ArrayList<>(allPositions);
		Collections.sort(positions, (a, b) -> {
			try {
				return Integer.compare(Integer.parseInt(a), Integer.parseInt(b));
			} catch (NumberFormatException e) {
				return a.compareTo(b);
			}
		});
		log.info("Position排序結果: " + String.join(", ", positions));

		// 7. 創建Position數據部分
		for (String position : positions) {
			log.info("處理Position: " + position);
			// 創建Position標題行
			Row positionRow = sheet.createRow(rowNum++);
			Cell positionCell = positionRow.createCell(0);
			positionCell.setCellValue("Position " + position);
			positionCell.setCellStyle(headerStyle);

			// 合併標題單元格
			for (int i = 1; i < columnCount; i++) {
				Cell cell = positionRow.createCell(i);
				cell.setCellStyle(headerStyle);
			}
			sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, columnCount - 1));

			// 創建數據行
			Row tcdRow = sheet.createRow(rowNum++);
			Cell tcdLabelCell = tcdRow.createCell(0);
			tcdLabelCell.setCellValue("TCD");
			tcdLabelCell.setCellStyle(headerStyle);

			Row bcdRow = sheet.createRow(rowNum++);
			Cell bcdLabelCell = bcdRow.createCell(0);
			bcdLabelCell.setCellValue("BCD");
			bcdLabelCell.setCellStyle(headerStyle);

			Row pshRow = sheet.createRow(rowNum++);
			Cell pshLabelCell = pshRow.createCell(0);
			pshLabelCell.setCellValue("PSH");
			pshLabelCell.setCellStyle(headerStyle);

			Row msRow = sheet.createRow(rowNum++);
			Cell msLabelCell = msRow.createCell(0);
			msLabelCell.setCellValue("M-S");
			msLabelCell.setCellStyle(blueStyle);

			// 為每個HT-DOM組填充數據
			for (int i = 0; i < htDomKeys.size(); i++) {
				String htDomKey = htDomKeys.get(i);
				Map<String, Map<String, String>> positionDataMap = htDomData.getOrDefault(htDomKey, new HashMap<>());
				Map<String, String> dataMap = positionDataMap.getOrDefault(position, new HashMap<>());
				log.info("  HT-DOM鍵: " + htDomKey + ", Position: " + position + ", 找到數據: " + !dataMap.isEmpty());

				if (!dataMap.isEmpty()) {
					log.info("    數據內容: measureId=" + dataMap.getOrDefault("measure_id", "未知") + ", TCD DX="
							+ dataMap.getOrDefault("TCD DX-95%", "無") + ", TCD DY="
							+ dataMap.getOrDefault("TCD DY", "無") + ", PS-BOT-DX="
							+ dataMap.getOrDefault("PS-BOT-DX", "無") + ", PS-BOT-DY="
							+ dataMap.getOrDefault("PS-BOT-DY", "無") + ", PS-Hp=" + dataMap.getOrDefault("PS-Hp", "無"));
				}
				int xCol = i * 2 + 1;
				int yCol = i * 2 + 2;

				if (!dataMap.isEmpty()) {
					// 獲取所有需要的數據
					String tcdX = dataMap.getOrDefault("TCD DX-95%", "");
					String tcdY = dataMap.getOrDefault("TCD DY", "");
					String bcdX = dataMap.getOrDefault("PS-BOT-DX", "");
					String bcdY = dataMap.getOrDefault("PS-BOT-DY", "");
					String psh = dataMap.getOrDefault("PS-Hp", "");

					log.info("    數據值: TCD DX=" + tcdX + ", TCD DY=" + tcdY + ", PS-BOT-DX=" + bcdX + ", PS-BOT-DY="
							+ bcdY + ", PS-Hp=" + psh);
					// 填充TCD數據
					setNumericCellValue(tcdRow.createCell(xCol), dataMap.getOrDefault("TCD DX-95%", ""), numberStyle);
					setNumericCellValue(tcdRow.createCell(yCol), dataMap.getOrDefault("TCD DY", ""), numberStyle);

					// 填充BCD數據
					setNumericCellValue(bcdRow.createCell(xCol), dataMap.getOrDefault("PS-BOT-DX", ""), numberStyle);
					setNumericCellValue(bcdRow.createCell(yCol), dataMap.getOrDefault("PS-BOT-DY", ""), numberStyle);

					// 填充PSH數據
					setNumericCellValue(pshRow.createCell(xCol), dataMap.getOrDefault("PS-Hp", ""), numberStyle);
					pshRow.createCell(yCol).setCellStyle(numberStyle); // Y列保持空白
				} else {
					log.info("    沒有找到數據");
					// 如果沒有數據，填充空單元格
					tcdRow.createCell(xCol).setCellStyle(numberStyle);
					tcdRow.createCell(yCol).setCellStyle(numberStyle);
					bcdRow.createCell(xCol).setCellStyle(numberStyle);
					bcdRow.createCell(yCol).setCellStyle(numberStyle);
					pshRow.createCell(xCol).setCellStyle(numberStyle);
					pshRow.createCell(yCol).setCellStyle(numberStyle);
				}

				// M-S行數據處理保持不變
				if (i == 0) {
					Cell dashCell = msRow.createCell(xCol);
					dashCell.setCellValue("-");
					dashCell.setCellStyle(basicStyle);

					Cell zeroCell = msRow.createCell(yCol);
					zeroCell.setCellValue(0.000);
					zeroCell.setCellStyle(numberStyle);
				} else {
					msRow.createCell(xCol).setCellStyle(basicStyle);
					msRow.createCell(yCol).setCellStyle(basicStyle);
				}
			}
		}

		return rowNum;
	}

	// 為單個SlideID生成完整報表
	private int createSlideReport(XSSFSheet sheet, int rowNum, String slideId, List<String> measureIds,
			Map<String, Map<String, String>> attributesByMeasure, Map<String, Map<String, String>> htDomByMeasure,
			Map<String, Map<String, List<String>>> positionsData, CellStyle headerStyle, CellStyle basicStyle,
			CellStyle numberStyle, CellStyle orangeStyle, CellStyle blueStyle) {

		int startRow = rowNum;

		// 1. 為每個measure_id建立HT-DOM鍵並分組
		Map<String, List<String>> htDomGroups = new LinkedHashMap<>();

		for (String measureId : measureIds) {
			Map<String, String> htDomValues = htDomByMeasure.getOrDefault(measureId, new HashMap<>());
			String ht = htDomValues.getOrDefault("HT", "100.0");
			String dom = htDomValues.getOrDefault("DOM", "");

			// 建立唯一的HT-DOM鍵
			String htDomKey = ht + "-" + dom;

			// 將measure_id加入對應的HT-DOM組
			htDomGroups.computeIfAbsent(htDomKey, k -> new ArrayList<>()).add(measureId);
		}

		// 獲取唯一的HT-DOM組
		List<String> uniqueHtDomKeys = new ArrayList<>(htDomGroups.keySet());

		// 計算真正需要的列數
		int columnCount = uniqueHtDomKeys.size() * 2 + 1; // 每個HT-DOM組佔2列(X和Y)

		// 2. 創建Item行
		Row itemRow = sheet.createRow(rowNum++);
		Cell itemLabelCell = itemRow.createCell(0);
		itemLabelCell.setCellValue("Item");
		itemLabelCell.setCellStyle(headerStyle);

		Cell itemValueCell = itemRow.createCell(1);
		itemValueCell.setCellValue(slideId);
		itemValueCell.setCellStyle(basicStyle);

		// 跨列合併Item值
		if (uniqueHtDomKeys.size() > 1) {
			sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 1, columnCount - 1));
			// 設置每個格子的樣式
			for (int i = 2; i < columnCount; i++) {
				Cell cell = itemRow.createCell(i);
				cell.setCellStyle(basicStyle);
			}
		}

		// 3. 創建動態屬性行
		if (!measureIds.isEmpty()) {
			// 獲取第一個measureId的所有屬性作為屬性名列表
			String firstMeasureId = measureIds.get(0);
			Map<String, String> firstMeasureAttrs = attributesByMeasure.getOrDefault(firstMeasureId, new HashMap<>());

			// 為每個屬性創建一行
			for (String attrName : firstMeasureAttrs.keySet()) {
				Row attrRow = sheet.createRow(rowNum++);
				Cell attrLabelCell = attrRow.createCell(0);
				attrLabelCell.setCellValue(attrName);
				attrLabelCell.setCellStyle(headerStyle);

				// 為每個HT-DOM組填充屬性值
				for (int i = 0; i < uniqueHtDomKeys.size(); i++) {
					String htDomKey = uniqueHtDomKeys.get(i);
					List<String> groupMeasureIds = htDomGroups.get(htDomKey);

					// 使用組內第一個measureId的屬性值
					if (!groupMeasureIds.isEmpty()) {
						String measureId = groupMeasureIds.get(0);
						Map<String, String> attrs = attributesByMeasure.getOrDefault(measureId, new HashMap<>());
						String attrValue = attrs.getOrDefault(attrName, "");

						// 選擇樣式（Mask使用藍色）
						CellStyle valueStyle = attrName.equals("Mask") ? blueStyle : basicStyle;

						// 填充值（每個HT-DOM組佔用2列）
						Cell attrValueCell = attrRow.createCell(i * 2 + 1);
						attrValueCell.setCellValue(attrValue);
						attrValueCell.setCellStyle(valueStyle);

						// 跨兩列合併單元格
						sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, i * 2 + 1, i * 2 + 2));

						// 設置第二列單元格樣式
						Cell mergedCell = attrRow.createCell(i * 2 + 2);
						mergedCell.setCellStyle(valueStyle);
					}
				}
			}
		}

		// 4. 創建HT (%)行
		Row htRow = sheet.createRow(rowNum++);
		Cell htLabelCell = htRow.createCell(0);
		htLabelCell.setCellValue("HT (%)");
		htLabelCell.setCellStyle(headerStyle);

		// 5. 創建Dom行
		Row domRow = sheet.createRow(rowNum++);
		Cell domLabelCell = domRow.createCell(0);
		domLabelCell.setCellValue("Dom");
		domLabelCell.setCellStyle(headerStyle);

		log.info("在 SlideID: " + slideId + " 中填充 HT-DOM 數據");
		log.info("唯一的 HT-DOM 鍵: " + String.join(", ", uniqueHtDomKeys));

		// 填充HT和Dom數據
		for (int i = 0; i < uniqueHtDomKeys.size(); i++) {
			String htDomKey = uniqueHtDomKeys.get(i);
			String[] parts = htDomKey.split("-");
			String ht = parts[0];
			String dom = parts.length > 1 ? parts[1] : "";

			// HT值
			Cell htValueCell = htRow.createCell(i * 2 + 1);
			htValueCell.setCellValue(ht);
			htValueCell.setCellStyle(orangeStyle);

			sheet.addMergedRegion(new CellRangeAddress(rowNum - 2, rowNum - 2, i * 2 + 1, i * 2 + 2));
			Cell htMergedCell = htRow.createCell(i * 2 + 2);
			htMergedCell.setCellStyle(orangeStyle);

			// Dom值
			Cell domValueCell = domRow.createCell(i * 2 + 1);
			domValueCell.setCellValue(dom);
			domValueCell.setCellStyle(orangeStyle);

			sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, i * 2 + 1, i * 2 + 2));
			Cell domMergedCell = domRow.createCell(i * 2 + 2);
			domMergedCell.setCellStyle(orangeStyle);
		}

		// 6. 創建X/Y行
		Row xyRow = sheet.createRow(rowNum++);
		Cell emptyCellXY = xyRow.createCell(0);
		emptyCellXY.setCellStyle(headerStyle);

		for (int i = 0; i < uniqueHtDomKeys.size(); i++) {
			Cell xCell = xyRow.createCell(i * 2 + 1);
			xCell.setCellValue("X");
			xCell.setCellStyle(basicStyle);

			Cell yCell = xyRow.createCell(i * 2 + 2);
			yCell.setCellValue("Y");
			yCell.setCellStyle(basicStyle);
		}

		// 7. 創建Position數據部分
		List<String> positions = new ArrayList<>(positionsData.keySet());
		Collections.sort(positions, (a, b) -> {
			try {
				return Integer.compare(Integer.parseInt(a), Integer.parseInt(b));
			} catch (NumberFormatException e) {
				return a.compareTo(b);
			}
		});

		// 遍歷所有position
		for (String position : positions) {
			Map<String, List<String>> dataNameMap = positionsData.get(position);

			// 創建Position標題行
			Row positionRow = sheet.createRow(rowNum++);
			Cell positionCell = positionRow.createCell(0);
			positionCell.setCellValue("Position " + position);
			positionCell.setCellStyle(headerStyle);

			// 合併標題單元格
			for (int i = 1; i < columnCount; i++) {
				Cell cell = positionRow.createCell(i);
				cell.setCellStyle(headerStyle);
			}
			sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, columnCount - 1));

			// 創建數據行
			Row tcdRow = sheet.createRow(rowNum++);
			Cell tcdLabelCell = tcdRow.createCell(0);
			tcdLabelCell.setCellValue("TCD");
			tcdLabelCell.setCellStyle(headerStyle);

			Row bcdRow = sheet.createRow(rowNum++);
			Cell bcdLabelCell = bcdRow.createCell(0);
			bcdLabelCell.setCellValue("BCD");
			bcdLabelCell.setCellStyle(headerStyle);

			Row pshRow = sheet.createRow(rowNum++);
			Cell pshLabelCell = pshRow.createCell(0);
			pshLabelCell.setCellValue("PSH");
			pshLabelCell.setCellStyle(headerStyle);

			Row msRow = sheet.createRow(rowNum++);
			Cell msLabelCell = msRow.createCell(0);
			msLabelCell.setCellValue("M-S");
			msLabelCell.setCellStyle(blueStyle);

			// 獲取數據列表
			List<String> tcdXValues = dataNameMap.getOrDefault("TCD DX-95%", new ArrayList<>());
			List<String> tcdYValues = dataNameMap.getOrDefault("TCD DY", new ArrayList<>());
			List<String> bcdXValues = dataNameMap.getOrDefault("PS-BOT-DX", new ArrayList<>());
			List<String> bcdYValues = dataNameMap.getOrDefault("PS-BOT-DY", new ArrayList<>());
			List<String> pshValues = dataNameMap.getOrDefault("PS-Hp", new ArrayList<>());

			// 直接為每個HT-DOM組填充數據
			for (int i = 0; i < uniqueHtDomKeys.size(); i++) {
				String htDomKey = uniqueHtDomKeys.get(i);
				List<String> groupMeasureIds = htDomGroups.get(htDomKey);
				log.info("處理 HT-DOM 組: " + htDomKey + ", 包含 MeasureIDs: " + String.join(", ", groupMeasureIds));
				int xCol = i * 2 + 1;
				int yCol = i * 2 + 2;

				// 針對此HT-DOM組，查找與此position相關的measureId
				boolean dataFound = false;
				for (String measureId : groupMeasureIds) {
					// 嘗試從measureIds列表中找到此measureId的索引
					int measureIndex = measureIds.indexOf(measureId);

					// 檢查該measureId是否與當前position相關，並且確保數據索引有效
					if (measureIndex >= 0 && measureIndex < tcdXValues.size() && positionsData.containsKey(position)
							&& positionsData.get(position).containsKey("TCD DX-95%")) {

						// TCD行數據
						setNumericCellValue(tcdRow.createCell(xCol), tcdXValues.get(measureIndex), numberStyle);
						setNumericCellValue(tcdRow.createCell(yCol), tcdYValues.get(measureIndex), numberStyle);

						// BCD行數據
						setNumericCellValue(bcdRow.createCell(xCol), bcdXValues.get(measureIndex), numberStyle);
						setNumericCellValue(bcdRow.createCell(yCol), bcdYValues.get(measureIndex), numberStyle);

						// PSH行數據
						setNumericCellValue(pshRow.createCell(xCol), pshValues.get(measureIndex), numberStyle);
						pshRow.createCell(yCol).setCellStyle(numberStyle); // Y列保持空白

						dataFound = true;
						break;
					}
				}

				// 如果沒找到數據，填充空單元格
				if (!dataFound) {
					tcdRow.createCell(xCol).setCellStyle(numberStyle);
					tcdRow.createCell(yCol).setCellStyle(numberStyle);
					bcdRow.createCell(xCol).setCellStyle(numberStyle);
					bcdRow.createCell(yCol).setCellStyle(numberStyle);
					pshRow.createCell(xCol).setCellStyle(numberStyle);
					pshRow.createCell(yCol).setCellStyle(numberStyle);
				}

				// M-S行數據處理保持不變
				if (i == 0) {
					Cell dashCell = msRow.createCell(xCol);
					dashCell.setCellValue("-");
					dashCell.setCellStyle(basicStyle);

					Cell zeroCell = msRow.createCell(yCol);
					zeroCell.setCellValue(0.000);
					zeroCell.setCellStyle(numberStyle);
				} else {
					msRow.createCell(xCol).setCellStyle(basicStyle);
					msRow.createCell(yCol).setCellStyle(basicStyle);
				}
			}
		}

		return rowNum;
	}

	// 輔助方法設置數值單元格
	private void setNumericCellValue(Cell cell, String value, CellStyle style) {
		if (value != null && !value.isEmpty()) {
			try {
				cell.setCellValue(Double.parseDouble(value));
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

	// 新增的輔助方法用於創建數據行
	private void createDataRow(XSSFSheet sheet, int rowNum, String label, List<String> xValues, List<String> yValues,
			int measureCount, CellStyle numberStyle) {
		Row row = sheet.createRow(rowNum);
		row.createCell(0).setCellValue(label);
		row.getCell(0).setCellStyle(numberStyle);

		for (int i = 0; i < measureCount; i++) {
			int xCol = i * 2 + 1;
			int yCol = i * 2 + 2;

			// 處理X值
			if (xValues != null && i < xValues.size()) {
				Cell xCell = row.createCell(xCol);
				setNumericCellValue(xCell, xValues.get(i), numberStyle);
			} else {
				row.createCell(xCol).setCellStyle(numberStyle);
			}

			// 處理Y值
			if (yValues != null && i < yValues.size()) {
				Cell yCell = row.createCell(yCol);
				setNumericCellValue(yCell, yValues.get(i), numberStyle);
			} else {
				row.createCell(yCol).setCellStyle(numberStyle);
			}
		}
	}

	// 輔助方法設置數值單元格

	// 創建數值單元格
	private void createValueCell(Row row, int column, String value, CellStyle style) {
		Cell cell = row.createCell(column);

		if (value != null && !value.isEmpty()) {
			try {
				double numValue = Double.parseDouble(value);
				cell.setCellValue(numValue);
			} catch (NumberFormatException e) {
				cell.setCellValue(value);
			}
		} else {
			cell.setCellValue("");
		}

		cell.setCellStyle(style);
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
	private CellStyle createNumberStyle(XSSFWorkbook workbook) {
		CellStyle style = createBasicStyle(workbook);
		style.setDataFormat(workbook.createDataFormat().getFormat("#,##0.000"));
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

	// 用于排序measureId的辅助类
	private class MeasureInfo {
		String measureId;
		double htValue;
		String dom;

		MeasureInfo(String measureId, double htValue, String dom) {
			this.measureId = measureId;
			this.htValue = htValue;
			this.dom = dom;
		}
	}
}