package infoqoch.writer.concreteplain;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target( {ElementType.FIELD } )
@Retention(RetentionPolicy.RUNTIME)
public @interface PlainCellMapper {
	int order();
	String columnName();
	int columnWidth() default 5000;
}
