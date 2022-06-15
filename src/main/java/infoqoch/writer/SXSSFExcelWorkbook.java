package infoqoch.writer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/*
- 시트가 유일하며, 시트 당 하나의 DTO만을 사용할 경우, 하나의 클래스(구 ExcelWriter -> 현 ExcelWorkbook)로 해결 가능하였음. (최초 패키지)
- 시트가 다수일 경우, 하나의 클래스로 모든 타입의 시트를 담보하기에는 어려웠음. 제네릭은 가변 갯수에 대하여 반영할 수 없었음.
	- 그러니까, ExcelWriter의 생성자에 가변갯수의 클래스를 넣고, 해당 클래스를 배열로 보관을 할 경우, 리플렉션을 통해 해당 타입들의 정합성은 판단할 수 있었음.
	- 하지만 런타임 시점에서 삽입하는 인자의 정합성은 결국 캐스팅으로 해소할 수밖에 없었고, 이는 안정되지 않음.
- 시트 당 하나의 DTO만을 사용하기 때문에, 결과적으로 제네릭을 통한 컴파일 에러를 체크하기 위해서는, 시트에 따른 클래스를 구현해야 했음.
- 그리고 ExcelWriter는 factory 패턴을 활용하여 인자의 타입이 ExcelSheet인지만 판단하고, appendSheet 으로 시트를 생성하는 기능으로 한정하였음.
- ExcelSheet을 구현 과정에서 한 번 더 분리해야 했음.
	- Builder와 그 자신인데, 리플렉션은 일회만 하고 그것은 계속 반복하여 사용. 하지만 Dto로 받는 리스트는 계속 변화함.
	- 이 두 가지를 분리해야 했고, 전자는 DefaultExcelSheetBuilder이며 후자는 ExcelSheet을 구현한 DefaultExcelSheet 임.
	- 전자는 리플렉션 데이터를 가지고 있다. 자신의 타입을 가진 리스트를 값으로 받아, DefaultExcelSheet를 리턴함.
	- 후자는 전자의 리플렉션 데이터와 리스트를 가지고 generateSheet()을 구현한다.
- DefaultExcelSheetBuilder 또한 인터페이스를 통해 구현해야할까 싶었지만, 그렇게 하지는 않았다. 중요한 지점은 ExcelSheet을 구현한 것이지, 그것을 구현하기 위한 Builder를 강제할 이유는 없기 때문이다.
*/

public class SXSSFExcelWorkbook implements ExcelWorkbook {

	private final SXSSFWorkbook wb;
	private final ExcelStyleFactory styleFactory;

	public SXSSFExcelWorkbook() {
		wb = new SXSSFWorkbook(new XSSFWorkbook());
		styleFactory = new ExcelStyleFactory() {
			@Override
			public Font createFont() {
				return wb.createFont();
			}
			@Override
			public CellStyle createCellStyle() {
				return wb.createCellStyle();
			}
		};
	}

	@Override
	public SXSSFExcelWorkbook appendSheet(ExcelSheet<?> excelSheet) {
		excelSheet.generateSheet(wb.createSheet(), styleFactory);

//		고민지점. 예외가 발생하면 시트를 그냥 지워버리는 것이 맞을까? 그런데 이 과정이 없으면 시트는 남아있는 상태가 된다. 음.
// 		일단은 복구는 클라이언트가 알아서 한다는 마인드로 냅두자. 굳이 지우지 말자.
//		Sheet sheet = wb.createSheet();
//		try {
//			excelSheet.generateSheet(wb.createSheet(), styleFactory);
//		}catch (Exception e) {
//			wb.removeSheetAt(wb.getSheetIndex(sheet));
//			throw e;
//		}

		return this;
	}

	public void write(File file) throws IOException {
		try(FileOutputStream fos = new FileOutputStream(file)){
			write(fos);
		}
	}

	@Override
	public void write(OutputStream stream) throws IOException {
		wb.write(stream);
	}
}

