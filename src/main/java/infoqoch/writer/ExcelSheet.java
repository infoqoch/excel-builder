package infoqoch.writer;

import org.apache.poi.ss.usermodel.Sheet;

public interface ExcelSheet<E> {
	void generateSheet(Sheet sheet, ExcelStyleFactory styleFactory);

	// 기존의 형태 sheet만 넘기는 것은 없앴음.
	// 스타일을 넣고 넣지 않고를 메서드로 판단하지 않고, CellStyle 을 매개변수로 무조건 제공한 후, 사용 여부를 결정하는 방식으로 진행.
	// void generateSheet(Sheet sheet);

	// 아래의 코드를 동작시킬 경우 문제가 발생. style은 하나이며, 해당 스타일이 중복되어 사용될 가능성이 매우 높음.
	// 이러한 방식보다는 특정 스타일이 적용되지 않는 스타일을 객체를 새로 생성해서 전달하는 방식으로 구현함.
	// void generateStyleSheet(Sheet sheet, CellStyle style, Font font) ;
}
