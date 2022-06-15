package infoqoch.writer;

import java.io.IOException;
import java.io.OutputStream;

public interface ExcelWorkbook {
	ExcelWorkbook appendSheet(ExcelSheet<?> excelSheet) throws IllegalAccessException;

	void write(OutputStream stream) throws IOException ;
}

