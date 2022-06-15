package infoqoch.writer.concretestyle;

import infoqoch.writer.ExcelSheet;
import infoqoch.writer.ExcelStyleFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;
import java.util.Map;

public class StyleExcelSheet<E> implements ExcelSheet<E> {
	private final Map<Integer, StyleCellInfo> cellInfo;
	private final List<E> list;

	StyleExcelSheet(Map<Integer, StyleCellInfo> cellInfo, List<E> list) {
		this.cellInfo = cellInfo;
		this.list = list;
	}

	private CellStyle applyStyles(StyleCellInfo info, ExcelStyleFactory styler) {
		CellStyle cellStyle = styler.createCellStyle();
		info.applyStyles(cellStyle, styler.createFont());
		return cellStyle;
	}

	@Override
	public void generateSheet(Sheet sheet, ExcelStyleFactory styler) {
		int headRows = generateHead(list.get(0).getClass(), sheet, styler);
		generateBody(list, sheet, headRows);
	}

	private int generateHead(Class<?> clazz, Sheet sheet, ExcelStyleFactory styler) {
		int headRows = 0;
		Row objfirstRow = sheet.createRow(headRows++);
		cellInfo.keySet().forEach(key -> {
			Cell cell = objfirstRow.createCell(key);
			cell.setCellValue(cellInfo.get(key).getColumnName());
			cell.setCellStyle(applyStyles(cellInfo.get(key), styler));
		});
		return headRows;
	}

	private void generateBody(List<?> list, Sheet sheet, int headRows) {
		for (int i = 0; i < list.size(); i++) {
			Row objRow = sheet.createRow(i + headRows);
			Row headRow = sheet.getRow(0);

			for (int j = 0; j < cellInfo.size(); j++) {
				Cell cell = objRow.createCell(j);
				cell.setCellValue(String.valueOf(cellInfo.get(j).getValue(list.get(i))));
				if(!cellInfo.get(j).isApplyStyleOnlyHeader()) {
					cell.setCellStyle(headRow.getCell(j).getCellStyle());
				}
			}
		}
	}
}