package infoqoch.writer.concretestyle;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

public class StyleExcelSheetBuilder<E> {
	private final Map<Integer, StyleCellInfo> sheetInfo;

	public StyleExcelSheetBuilder(Class<E> clazz) {
		sheetInfo = generateSheetInfo(clazz);
	}

	public StyleExcelSheet<E> getInstance(List<E> list){
		return new StyleExcelSheet<>(sheetInfo, list);
	}

	private static Map<Integer, StyleCellInfo> generateSheetInfo(Class<?> clazz) {

		Map<Integer, StyleCellInfo> target = extractSheetInfoWithReflection(clazz);

	    validOrdered(target);

	    return target;
	}

	private static Map<Integer, StyleCellInfo> extractSheetInfoWithReflection(Class<?> clazz) {
		Map<Integer, StyleCellInfo> target = new ConcurrentHashMap<>();

		Field[] declaredField = clazz.getDeclaredFields();
	    for(Field field : declaredField) {
	        if(field.isAnnotationPresent(StyleCellMapper.class)) {
	        	// non-static 이너클래스는 자신을 감싸는 인스턴스를 참조한다. 만약 존재하면 생략한다. 가장 좋은 방법은 static class로 이너 클래스를 정의한다.
	        	if(field.getName().startsWith("this$"))
	        		continue;

        		StyleCellMapper excelWriterInfo = field.getAnnotation(StyleCellMapper.class);
	        	if(excelWriterInfo.order()<0) {
	        		throw new IllegalStateException("order should be greater than or equal to 0");
	        	}
	        	if(target.containsKey(excelWriterInfo.order()))
	        		throw new IllegalStateException("not allowed duplicated order");
	        	if(excelWriterInfo.columnName()==null) {
	        		throw new IllegalStateException("columnName should not be null");
	        	}
	        	if(excelWriterInfo.columnName().trim().length()==0) {
	        		throw new IllegalStateException("columnName should not be empty");
	        	}

	        	field.setAccessible(true);
	        	target.put(excelWriterInfo.order(), new StyleCellInfo(excelWriterInfo, field));
	        }else {
	        	throw new IllegalStateException("All field should be annotated with ExcelWriterInfo");
	        }
	    }
		return target;
	}

	private static void validOrdered(Map<Integer, StyleCellInfo> target) {
		List<Integer> keys = target.keySet().stream().sorted().collect(Collectors.toList());

		if(!keys.get(0).equals(0))
			throw new IllegalStateException("order should be start with 0");

	    for(int i=0; i<keys.size(); i++) {
	    	if(keys.get(i)-i!=0)
	    		throw new IllegalStateException("order should be increased the value of a variable by 1");
		}
	}
}
