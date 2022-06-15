package infoqoch.writer.concreteplain;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

public class PlainExcelSheetBuilder<E> {
	private final Map<Integer, PlainCellInfo> sheetInfo;

	public PlainExcelSheetBuilder(Class<E> clazz) {
		sheetInfo = generateSheetInfo(clazz);
	}

	public PlainExcelSheet<E> getInstance(List<E> list){
		return new PlainExcelSheet<>(sheetInfo, list);
	}

	private static Map<Integer, PlainCellInfo> generateSheetInfo(Class<?> clazz) {

		Map<Integer, PlainCellInfo> target = extractSheetInfoWithReflection(clazz);

	    validOrdered(target);

	    return target;
	}

	private static Map<Integer, PlainCellInfo> extractSheetInfoWithReflection(Class<?> clazz) {
		Map<Integer, PlainCellInfo> target = new ConcurrentHashMap<>();

		Field[] declaredField = clazz.getDeclaredFields();
	    for(Field field : declaredField) {
	        if(field.isAnnotationPresent(PlainCellMapper.class)) {
	        	// non-static 이너클래스는 자신을 감싸는 인스턴스를 참조한다. 만약 존재하면 생략한다. 가장 좋은 방법은 static class로 이너 클래스를 정의한다.
	        	if(field.getName().startsWith("this$"))
	        		continue;

        		PlainCellMapper excelWriterInfo = field.getAnnotation(PlainCellMapper.class);
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
	        	target.put(excelWriterInfo.order(), new PlainCellInfo(excelWriterInfo.order(), excelWriterInfo.columnName(), field));
	        }else {
	        	throw new IllegalStateException("All field should be annotated with ExcelWriterInfo");
	        }
	    }
		return target;
	}

	private static void validOrdered(Map<Integer, PlainCellInfo> target) {
		List<Integer> keys = target.keySet().stream().sorted().collect(Collectors.toList());

		if(!keys.get(0).equals(0))
			throw new IllegalStateException("order should be start with 0");

	    for(int i=0; i<keys.size(); i++) {
	    	if(keys.get(i)-i!=0)
	    		throw new IllegalStateException("order should be increased the value of a variable by 1");
		}
	}
}
