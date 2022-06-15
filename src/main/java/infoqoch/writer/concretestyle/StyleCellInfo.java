package infoqoch.writer.concretestyle;

import java.lang.reflect.Field;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter(value = AccessLevel.PRIVATE)
@AllArgsConstructor
class StyleCellInfo{
	private final StyleCellMapper mapper;
	private final Field field;

	public String getColumnName() {
		return getMapper().columnName();
	}

	public Object getValue(Object obj) {
		try {
			return getField().get(obj);
		}catch (Exception e) {
			throw new IllegalStateException("Cannot not get field with reflection. caused message : " + e.getMessage());
		}
	}

	public boolean isApplyStyleOnlyHeader() {
		return mapper.applyStyleOnlyHeader();
	}

	public void applyStyles(CellStyle cellStyle, Font font) {
		applyStyle(cellStyle);
		applyFont(cellStyle, font);
	}

	private void applyStyle(CellStyle cellStyle) {
		if(getMapper().cellColor().equals(IndexedColors.WHITE))
			return;
		cellStyle.setFillForegroundColor(getMapper().cellColor().getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}

	private void applyFont(CellStyle cellStyle, Font font) {
		if(getMapper().fontStyle().equals(FontStyle.NONE))
			return;
		getMapper().fontStyle().applyFont(font);
		cellStyle.setFont(font);
	}
}