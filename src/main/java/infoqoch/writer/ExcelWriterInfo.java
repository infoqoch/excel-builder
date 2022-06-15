package infoqoch.writer;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.IndexedColors;

@Target( {ElementType.FIELD } )
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelWriterInfo {
	int order();
	String columnName();
	int columnWidth() default 5000;

	FontStyle fontStyle() default FontStyle.NONE;
	boolean applyFontStyleOnlyHeader() default true;

	IndexedColors cellColor() default IndexedColors.BLACK;
	boolean applyCellColorOnlyHeader() default true;

	enum FontStyle {
		NONE, BOLD, ITALIC
	}
}
