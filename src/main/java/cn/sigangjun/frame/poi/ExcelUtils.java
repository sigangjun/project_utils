package cn.sigangjun.frame.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.sigangjun.frame.poi.test.DeviceDto;

/**
 * 功能: POI实现把Excel转换成对应的Dto
 */
public class ExcelUtils {

	// 正则表达式 用于匹配属性的第一个字母
	private static final String REGEX = "[a-zA-Z]";

	/**
	 * 功能: Excel数据导入到数据库 参数: filePath[Excel表的所在路径] 参数: startRow[从第几行开始] 参数: endRow[到第几行结束 (0表示所有行; 正数表示到第几行结束; 负数表示到倒数第几行结束)] 参数: clazz[要返回的对象集合的类型]
	 */
	public static <T> List<T> importExcel(String filePath, Class<T> clazz) throws IOException {
		// 是否打印提示信息
		return doImportExcel(filePath, 0, 0, clazz);
	}

	/**
	 * 功能: Excel数据导入到数据库
	 * 
	 * @param filePath
	 *            Excel表的所在路径
	 * @param startRow
	 *            从第几行开始
	 * @param endRow
	 *            到第几行结束 (0表示所有行; 正数表示到第几行结束; 负数表示到倒数第几行结束)
	 * @param clazz
	 *            要返回的对象集合的类型
	 * @return
	 * @throws IOException
	 */
	public static <T> List<T> importExcel(String filePath, int startRow, int endRow, Class<T> clazz) throws IOException {
		// 是否打印提示信息
		return doImportExcel(filePath, startRow, endRow, clazz);
	}

	/**
	 * 功能:真正实现导入
	 */
	@SuppressWarnings("resource")
	private static <T> List<T> doImportExcel(String filePath, int startRow, int endRow, Class<T> clazz) throws IOException {
		// 判断文件是否存在
		File file = new File(filePath);
		if (!file.exists()) {
			throw new IOException("文件名为" + file.getName() + "Excel文件不存在！");
		}
		Workbook wb = null;
		List<Row> rowList = new ArrayList<Row>();
		ArrayList<String> columns = new ArrayList<>();
		FileInputStream fileInputStream = new FileInputStream(file);
		try {
			// 去读Excel
			if (filePath.endsWith("xls")) {
				wb = new HSSFWorkbook(fileInputStream);
			} else if (filePath.endsWith("xlsx")) {
				wb = new XSSFWorkbook(fileInputStream);
			} else {
				throw new RuntimeException("当前文件不是excel文件:" + filePath);
			}

			Sheet sheet = wb.getSheetAt(0);
			// 获取最后行号
			int lastRowNum = sheet.getLastRowNum();
			Row row = null;
			// 循环读取
			for (int i = startRow; i <= lastRowNum + endRow; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					if (i == startRow) {
						for (int j = 0; j < row.getLastCellNum(); j++) {
							String value = getCellValue(row.getCell(j));
							columns.add(value.trim());
						}
					} else {
						rowList.add(row);
					}
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (wb != null) {
				wb.close();
			}
			if (fileInputStream != null) {
				fileInputStream.close();
			}
		}
		return returnObjectList(rowList, clazz, columns);
	}

	/**
	 * 功能:获取单元格的值
	 */
	private static String getCellValue(Cell cell) {
		Object result = "";
		if (cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				result = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				result = cell.getNumericCellValue();
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				result = cell.getBooleanCellValue();
				break;
			case Cell.CELL_TYPE_FORMULA:
				result = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_ERROR:
				result = cell.getErrorCellValue();
				break;
			case Cell.CELL_TYPE_BLANK:
				break;
			default:
				break;
			}
		}
		return result.toString();
	}

	/**
	 * 功能:返回指定的对象集合
	 */
	private static <T> List<T> returnObjectList(List<Row> rowList, Class<T> clazz, List<String> columns) {
		List<T> objectList = null;
		T obj = null;
		String attribute = null;
		String attribute_en = null;
		String value = null;
		try {
			objectList = new ArrayList<T>();
			Field[] declaredFields = clazz.getDeclaredFields();
			for (Row row : rowList) {
				obj = (T) clazz.newInstance();
				for (Field field : declaredFields) {
					attribute = field.getName().toString();
					attribute_en = attribute;
					if (field.isAnnotationPresent(ExcelTitleAnnotation.class)) {
						ExcelTitleAnnotation annotation = field.getAnnotation(ExcelTitleAnnotation.class);
						attribute = annotation.value();
					}
					int j = columns.indexOf(attribute);
					value = getCellValue(row.getCell(j));
					setAttrributeValue(obj, attribute_en, value);
				}
				objectList.add(obj);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return objectList;
	}

	/**
	 * 功能:给指定对象的指定属性赋值
	 */
	private static void setAttrributeValue(Object obj, String attribute, String value) {
		// 得到该属性的set方法名
		String method_name = convertToMethodName(attribute, obj.getClass(), true);
		Method[] methods = obj.getClass().getMethods();
		for (Method method : methods) {
			/**
			 * 因为这里只是调用bean中属性的set方法，属性名称不能重复 所以set方法也不会重复，所以就直接用方法名称去锁定一个方法 （注：在java中，锁定一个方法的条件是方法名及参数）
			 */
			if (method.getName().equals(method_name)) {
				Class<?>[] parameterC = method.getParameterTypes();
				try {
					/**
					 * 如果是(整型,浮点型,布尔型,字节型,时间类型), 按照各自的规则把value值转换成各自的类型 否则一律按类型强制转换(比如:String类型)
					 */
					if (parameterC[0] == int.class || parameterC[0] == java.lang.Integer.class) {
						value = value.substring(0, value.lastIndexOf("."));
						method.invoke(obj, Integer.valueOf(value));
						break;
					} else if (parameterC[0] == float.class || parameterC[0] == java.lang.Float.class) {
						method.invoke(obj, Float.valueOf(value));
						break;
					} else if (parameterC[0] == double.class || parameterC[0] == java.lang.Double.class) {
						method.invoke(obj, Double.valueOf(value));
						break;
					} else if (parameterC[0] == byte.class || parameterC[0] == java.lang.Byte.class) {
						method.invoke(obj, Byte.valueOf(value));
						break;
					} else if (parameterC[0] == boolean.class || parameterC[0] == java.lang.Boolean.class) {
						method.invoke(obj, Boolean.valueOf(value));
						break;
					} else if (parameterC[0] == java.util.Date.class) {
						SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
						Date date = null;
						try {
							date = sdf.parse(value);
						} catch (Exception e) {
							e.printStackTrace();
						}
						method.invoke(obj, date);
						break;
					} else {
						method.invoke(obj, parameterC[0].cast(value));
						break;
					}
				} catch (IllegalArgumentException e) {
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					e.printStackTrace();
				} catch (SecurityException e) {
					e.printStackTrace();
				}
			}
		}
	}

	/**
	 * 功能:根据属性生成对应的set/get方法
	 */
	private static String convertToMethodName(String attribute, Class<?> objClass, boolean isSet) {
		/** 通过正则表达式来匹配第一个字符 **/
		Pattern p = Pattern.compile(REGEX);
		Matcher m = p.matcher(attribute);
		StringBuilder sb = new StringBuilder();
		/** 如果是set方法名称 **/
		if (isSet) {
			sb.append("set");
		} else {
			/** get方法名称 **/
			try {
				Field attributeField = objClass.getDeclaredField(attribute);
				/** 如果类型为boolean **/
				if (attributeField.getType() == boolean.class || attributeField.getType() == Boolean.class) {
					sb.append("is");
				} else {
					sb.append("get");
				}
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
		}
		/** 针对以下划线开头的属性 **/
		if (attribute.charAt(0) != '_' && m.find()) {
			sb.append(m.replaceFirst(m.group().toUpperCase()));
		} else {
			sb.append(attribute);
		}
		return sb.toString();
	}

	public static void main(String[] args) throws IOException {
		// ArrayList<String> array = new ArrayList<>();
		// array.add("aaa");
		// array.add("bbb");
		// array.add("ccc");
		// array.add("ddd");
		// int indexOf = array.indexOf("ccc");
		// System.out.println(indexOf);
		// if(1==1)return ;

		String filePath = "T:\\template.xlsx";

		int startRow = 0;
		int endRow = 0;
		List<DeviceDto> importExcel = ExcelUtils.importExcel(filePath, startRow, endRow, DeviceDto.class);
		for (DeviceDto dto : importExcel) {
			System.out.println(dto);
		}

	}

}