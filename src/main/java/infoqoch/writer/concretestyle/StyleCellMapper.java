package infoqoch.writer.concretestyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.IndexedColors;

@Target( {ElementType.FIELD } )
@Retention(RetentionPolicy.RUNTIME)
public @interface StyleCellMapper {
	// 기본
	int order();
	String columnName();
	int columnWidth() default 5000;


	boolean applyStyleOnlyHeader() default true;

	// 폰트스타일
	FontStyle fontStyle() default FontStyle.NONE;
	IndexedColors cellColor() default IndexedColors.WHITE;

}
