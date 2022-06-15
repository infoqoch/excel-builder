package infoqoch.writer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

// 기존 구현은 sheet에 의존하도록 하였다.
// 하지만 CellStyle이나 Font는 Workbook 에서만 객체를 가져올 수 있다.
// 그러므로 ExcelWriter에서 wb를 통해 아래의 코드를 구현한 후, sheet를 생성할 때 사용하도록 매개변수에 주입한다.
public interface ExcelStyleFactory {
	CellStyle createCellStyle();
	Font createFont();
}
