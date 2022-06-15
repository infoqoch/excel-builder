package infoqoch.writer.concreteplain;

import infoqoch.writer.SXSSFExcelWorkbook;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.assertj.core.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

public class PlainExcelWorkbookTest {

	PlainExcelSheetBuilder<ExcelTestDTO1> builder1 = new PlainExcelSheetBuilder<>(ExcelTestDTO1.class);
	PlainExcelSheetBuilder<ExcelTestDTO2> builder2 = new PlainExcelSheetBuilder<>(ExcelTestDTO2.class);

	@Test
	void basic_test() throws IOException {
		PlainExcelSheet<ExcelTestDTO1> sheet1 = builder1.getInstance(generateExcelTestDTO1(10));
		SXSSFExcelWorkbook wb = new SXSSFExcelWorkbook();
		wb.appendSheet(sheet1);
		File file = new File("C:\\data\\targetP1.xlsx");
		wb.write(file);
	}

//	@Test
	void twice_write_exception() throws IOException {
		PlainExcelSheet<ExcelTestDTO1> sheet1 = builder1.getInstance(generateExcelTestDTO1(10));
		SXSSFExcelWorkbook wb = new SXSSFExcelWorkbook();
		wb.appendSheet(sheet1);
		File file = new File("C:\\data\\targetP2.xlsx");
		wb.write(file);

		PlainExcelSheet<ExcelTestDTO2> sheet2 = builder2.getInstance(generateExcelTestDTO2(10));
		wb.appendSheet(sheet2);

		Assertions.assertThatThrownBy(() -> wb.write(file)).isInstanceOf(IllegalStateException.class);
	}

	@Data
	@AllArgsConstructor
	static class ExcelTestDTO1 {
		@PlainCellMapper(order = 0, columnName = "name")
		private String name;
		@PlainCellMapper(order = 1, columnName = "id")
		private int id;
	}

	@Data
	@AllArgsConstructor
	static class ExcelTestDTO2 {
		@PlainCellMapper(order = 0, columnName = "date")
		private LocalDateTime now;
		@PlainCellMapper(order = 1, columnName = "장소")
		private String place;
	}

	private List<ExcelTestDTO1> generateExcelTestDTO1(int size) {
		List<ExcelTestDTO1> list = new ArrayList<>();
		for (int i = 0; i < size; i++) {
			list.add(new ExcelTestDTO1(UUID.randomUUID().toString(), i));
		}
		return list;
	}

	private List<ExcelTestDTO2> generateExcelTestDTO2(int size) {
		List<ExcelTestDTO2> list = new ArrayList<>();

		LocalDateTime now = LocalDateTime.now();
		for (int i = 0; i < size; i++) {
			list.add(new ExcelTestDTO2(now.minusDays(i), "seoul" + i));
		}
		return list;
	}
}