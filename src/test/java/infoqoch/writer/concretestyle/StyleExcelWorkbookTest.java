package infoqoch.writer.concretestyle;

import infoqoch.writer.ExcelSheet;
import infoqoch.writer.ExcelStyleFactory;
import infoqoch.writer.SXSSFExcelWorkbook;
import infoqoch.writer.concreteplain.PlainCellMapper;
import infoqoch.writer.concreteplain.PlainExcelSheet;
import infoqoch.writer.concreteplain.PlainExcelSheetBuilder;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.assertj.core.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;


public class StyleExcelWorkbookTest {

	StyleExcelSheetBuilder<ExcelTestDTO> builder1 = new StyleExcelSheetBuilder<>(ExcelTestDTO.class);
	StyleExcelSheetBuilder<ExcelTestDTO2> builder2 = new StyleExcelSheetBuilder<>(ExcelTestDTO2.class);
	PlainExcelSheetBuilder<ExcelTestDTO3> builder3 = new PlainExcelSheetBuilder<>(ExcelTestDTO3.class);

	@Test
	@DisplayName("정상 동작. 가능한 기능 전체 사용")
	void basic_test() throws IOException {
		StyleExcelSheet<ExcelTestDTO> sheet1 = builder1.getInstance(generateExcelTestDTO(10));
		StyleExcelSheet<ExcelTestDTO2> sheet2 = builder2.getInstance(generateExcelTestDTO2(10));
		PlainExcelSheet<ExcelTestDTO3> sheet3  = builder3.getInstance(generateExcelTestDTO3(10));
		SXSSFExcelWorkbook wb = new SXSSFExcelWorkbook();
		wb.appendSheet(sheet1);
		wb.appendSheet(sheet2);
		wb.appendSheet(new ExcelSheet<Object>() {
			@Override
			public void generateSheet(Sheet sheet, ExcelStyleFactory styleFactory) {
				Row row = sheet.createRow(0);
				row.createCell(0).setCellValue("하이!!");
				CellStyle style = styleFactory.createCellStyle();
				style.setFillBackgroundColor(IndexedColors.BROWN.getIndex());
				style.setFillPattern(FillPatternType.BRICKS);
				row.setRowStyle(style);
				row.createCell(1).setCellValue("시작한다요!!!");
			}
		});
		wb.appendSheet(sheet3);

		File file = new File("C:\\data\\targetS1.xlsx");
		wb.write(file);
	}

	@Test
	@DisplayName("wb의 중복사용 허용")
	// wb의 중복 사용을 이전에는 하지 못하도록 하였는데, 굳이 그럴 필요가 없어 보인다. 그냥 사용. 알아서 하겠지!
	void reuse_okay() throws IOException {
		StyleExcelSheet<ExcelTestDTO> sheet1 = builder1.getInstance(generateExcelTestDTO(10));
		PlainExcelSheet<ExcelTestDTO3> sheet3  = builder3.getInstance(generateExcelTestDTO3(10));

		SXSSFExcelWorkbook wb = new SXSSFExcelWorkbook();
		wb.appendSheet(sheet1);
		File file = new File("C:\\data\\targetS21.xlsx");
		wb.write(file);

		wb.appendSheet(sheet3);
		File file2 = new File("C:\\data\\targetS22.xlsx");
		wb.write(file2);
	}

	@Test
	@DisplayName("sheet의 생성 과정에서 예외가 발생하더라도 이전 sheet 는 정상적으로 write 가능하다")
	// wb의 중복 사용을 이전에는 하지 못하도록 하였는데, 굳이 그럴 필요가 없어 보인다. 그냥 사용. 알아서 하겠지!
	void sheet_exception_recover_wb() throws IOException {
		StyleExcelSheet<ExcelTestDTO> sheet1 = builder1.getInstance(generateExcelTestDTO(10));
		StyleExcelSheet<ExcelTestDTO2> sheet2 = builder2.getInstance(generateExcelTestDTO2(10));



		SXSSFExcelWorkbook wb = new SXSSFExcelWorkbook();
		wb.appendSheet(sheet1);


		boolean isCatchException = false;
		try {
			wb.appendSheet(new ExcelSheet<Object>() {
				@Override
				public void generateSheet(Sheet sheet, ExcelStyleFactory styleFactory) {
					throw new IllegalArgumentException("예외가 터졌드아!");
				}
			});
		}catch (Exception e) {
			System.out.println("예외 발생!!");
			e.printStackTrace();
			isCatchException = true;
		}
		Assertions.assertThat(isCatchException).isTrue();

		File file = new File("C:\\data\\targetS3.xlsx");
		wb.write(file);
	}

	@Data
	@AllArgsConstructor
	static class ExcelTestDTO{
		@StyleCellMapper(order = 0, columnName = "name"
				, cellColor = IndexedColors.BLUE
				, applyStyleOnlyHeader = false)
		private String name;
		@StyleCellMapper(order = 1, columnName = "id"
				, cellColor = IndexedColors.BROWN)
		private int id;
	}

	@Data
	@AllArgsConstructor
	static class ExcelTestDTO2{
		@StyleCellMapper(order = 0, columnName = "date")
		private LocalDateTime now;
		@StyleCellMapper(order = 1, columnName = "그녀")
		private String her;
	}

	private List<ExcelTestDTO> generateExcelTestDTO(int size) {
		List<ExcelTestDTO> list = new ArrayList<>();
		for(int i=0; i<size; i++) {
			list.add(new ExcelTestDTO(UUID.randomUUID().toString(), i));
		}
		return list;
	}

	private List<ExcelTestDTO2> generateExcelTestDTO2(int size) {
		List<ExcelTestDTO2> list = new ArrayList<>();

		LocalDateTime now = LocalDateTime.now();
		for(int i=0; i<size; i++) {
			list.add(new ExcelTestDTO2(now.minusDays(i), "kim"+i));
		}
		return list;
	}


	@Data
	@AllArgsConstructor
	static class ExcelTestDTO3 {
		@PlainCellMapper(order = 1, columnName = "UUID")
		private String uuid;
		@PlainCellMapper(order = 0, columnName = "idx")
		private int idx;
	}

	private List<ExcelTestDTO3> generateExcelTestDTO3(int size) {
		List<ExcelTestDTO3> list = new ArrayList<>();
		for (int i = 0; i < size; i++) {
			list.add(new ExcelTestDTO3(UUID.randomUUID().toString(), i));
		}
		return list;
	}

}
