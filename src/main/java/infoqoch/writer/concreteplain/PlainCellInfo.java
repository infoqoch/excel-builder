package infoqoch.writer.concreteplain;

import java.lang.reflect.Field;

import lombok.Getter;
import lombok.ToString;

@Getter
@ToString
class PlainCellInfo{
	private final int order;
	private final String columnName;
	private final Field field;

	PlainCellInfo(int order, String columnName, Field field) {
		this.order = order;
		this.columnName = columnName;
		this.field = field;
	}
}