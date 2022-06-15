package infoqoch.writer.concreteplain;

import infoqoch.writer.ExcelSheet;
import infoqoch.writer.ExcelStyleFactory;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

public class PlainExcelSheet<E> implements ExcelSheet<E> {
	private final Map<Integer, PlainCellInfo> sheetInfo;
	private final List<E> list;

	PlainExcelSheet(Map<Integer, PlainCellInfo> sheetInfo, List<E> list) {
		this.sheetInfo = sheetInfo;
		this.list = list;
	}

	@Override
	public void generateSheet(Sheet sheet, ExcelStyleFactory styler){
		int headRows = generateHead(list.get(0).getClass(), sheet);
		generateBody(list, sheet, headRows);
	}

	private int generateHead(Class<?> clazz, Sheet sheet) {
		int headRows = 0;
		Row firstRow = sheet.createRow(headRows++);
		sheetInfo.keySet().forEach(key-> firstRow.createCell(key).setCellValue(sheetInfo.get(key).getColumnName()));
		return headRows;
	}

	private void generateBody(List<?> list, Sheet sheet, int headRows){
		for(int i=0; i<list.size(); i++) {
			Object dto = list.get(i);
			Row objRow = sheet.createRow(i+headRows);
			for(int j=0; j<sheetInfo.size(); j++) {
				Field field = sheetInfo.get(j).getField();
				field.setAccessible(true);
				try {
					Object value = field.get(dto);
					objRow.createCell(j).setCellValue(String.valueOf(value));
				}catch (Exception e) {
					throw new IllegalStateException("Cannot not get field with reflection. caused message : " + e.getMessage());
				}
			}
		}
	}
}
