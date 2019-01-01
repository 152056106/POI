package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * excel到处导入
 * 
 * @author LiMing
 * 
 */
public class GenerateExcel {


	/**
	 * 输出excel文件
	 * @param data 数据集合
	 * @param path 输出路径
	 */
	public static void outExcelFile(List<?> data, String path) {

		File file = new File(path);
		// 创建workbook
		HSSFWorkbook wb = new HSSFWorkbook();
		// 创建sheet
		Sheet sheet = wb.createSheet("sheel");
		// 创建表头行
		Row row = sheet.createRow(0);
		// 创建单元格样式
		HSSFCellStyle style = wb.createCellStyle();
		// 居中显示
		style.setAlignment(HorizontalAlignment.CENTER);
		/**
		 * 获取实体所有属性
		 * getFields()：获得某个类的所有的公共（public）的字段，包括父类中的字段。 
		 * getDeclaredFields()：获得某个类的所有声明的字段，即包括public、private和proteced，但是不包括父类的申明字段。
		 * 
		 * */
		Field[] fields = data.get(0).getClass().getDeclaredFields();
		// 列索引
		int index = 0;
		// 列名称
		String name = "";
		MyAnnotation myAnnotation;
		// 创建表头
		for (Field f : fields) {
			// 是否是注解
			if (f.isAnnotationPresent(MyAnnotation.class)) {
				// 获取注解
				myAnnotation = f.getAnnotation(MyAnnotation.class);
				// 获取列索引
				index = myAnnotation.columnIndex();
				// 列名称
				name = myAnnotation.columnName();
				// 创建单元格
				creCell(row, index, name, style);
			}
		}

		// 行索引  因为表头已经设置，索引行索引从1开始
		int rowIndex = 1;
		for (Object obj : data) {
			// 创建新行，索引加1,为创建下一行做准备
			row = sheet.createRow(rowIndex++);
			for (Field f : fields) {
				// 设置属性可访问
				f.setAccessible(true);
				// 判断是否是注解
				if (f.isAnnotationPresent(MyAnnotation.class)) {
					// 获取注解
					myAnnotation = f.getAnnotation(MyAnnotation.class);
					// 获取列索引
					index = myAnnotation.columnIndex();
					try {
						// 创建单元格     f.get(obj)从obj对象中获取值设置到单元格中
						creCell(row, index, String.valueOf(f.get(obj)), style);
					} catch (IllegalArgumentException e) {
						e.printStackTrace();
					} catch (IllegalAccessException e) {
						e.printStackTrace();
					}
				}
			}
		}

		FileOutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream(file);
			/**
			 * 写出数据,file 为路径
			 * */
			wb.write(outputStream);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			//释放资源
			try {
				if (wb != null) {
					try {
						wb.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
				if (outputStream != null) {
					outputStream.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/**
	 * 读取excel文件，并把读取到的数据封装到clazz中
	 * 
	 * @param path
	 *            文件路径
	 * @param clazz
	 *            实体类
	 * @return 返回clazz集合
	 */
	public static <T extends Object> List<T> readExcelFile(String path, Class<T> clazz) {
		// 存储excel数据
		List<T> list = new ArrayList<>();
		FileInputStream is = null;

		try {
			is = new FileInputStream(new File(path));
		} catch (FileNotFoundException e1) {
			throw new RuntimeException("文件路径异常");
		}

		Workbook wookbook = null;

		// 根据excel文件版本获取工作簿,判断excel 的文件类型
		if (path.endsWith(".xls")) {
			wookbook = xls(is);
		} else if (path.endsWith(".xlsx")) {
			wookbook = xlsx(is);
		} else {
			throw new RuntimeException("文件出错，非excel文件");
		}

		// 得到一个工作表
		Sheet sheet = wookbook.getSheetAt(0);

		// 获取行总数
		int rows = sheet.getLastRowNum() + 1;

		Row row;

		// 获取类所有属性
		Field[] fields = clazz.getDeclaredFields();

		T obj = null;
		int coumnIndex = 0;
		Cell cell = null;
		MyAnnotation myAnnotation = null;
		for (int i = 1; i < rows; i++) {
			// 获取excel行
			row = sheet.getRow(i);
			try {
				// 创建实体
				obj = clazz.newInstance();
				for (Field f : fields) {
					// 设置属性可访问
					f.setAccessible(true);
					// 判断是否是注解
					if (f.isAnnotationPresent(MyAnnotation.class)) {
						// 获取注解
						myAnnotation = f.getAnnotation(MyAnnotation.class);
						// 获取列索引
						coumnIndex = myAnnotation.columnIndex();
						// 获取单元格
						cell = row.getCell(coumnIndex);
						// 设置属性
						setFieldValue(obj, f, wookbook, cell);
					}
				}
				// 添加到集合中
				list.add(obj);
			} catch (InstantiationException e1) {
				e1.printStackTrace();
			} catch (IllegalAccessException e1) {
				e1.printStackTrace();
			}

		}

		try {
			//释放资源
			wookbook.close();
			is.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return list;
	}

	/**
	 * 设置属性值
	 * 
	 * @param obj
	 *            操作对象
	 * @param f
	 *            对象属性
	 * @param cell
	 *            excel单元格
	 */
	private static void setFieldValue(Object obj, Field f, Workbook wookbook, Cell cell) {
		try {
			if (f.getType() == int.class || f.getType() == Integer.class) {
				f.setInt(obj, getInt(cell));
			} else if (f.getType() == Double.class || f.getType() == double.class) {
				f.setDouble(obj, getDouble(null, cell));
			} else {
				f.set(obj, getString(cell));
			}
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 获取正数
	 * 
	 * @param cell
	 * @return
	 */
	@SuppressWarnings("deprecation")
	private static int getInt(Cell cell) {
		if (cell == null) {
			return 0;
		}
		if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
			return 0;
		}
		if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
			return Integer.parseInt(getString(cell));
		}
		return Integer.parseInt(NumberToTextConverter.toText(cell.getNumericCellValue()));
	}

	/**
	 * 获取double
	 * 
	 * @param cell
	 * @return
	 */
	@SuppressWarnings("deprecation")
	private static double getDouble(Workbook wookbook, Cell cell) {
		double d = 0;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			// 空白字符 返回0
			d = 0;
			break;
		case Cell.CELL_TYPE_FORMULA:
			// 公式
			FormulaEvaluator formulaEval = wookbook.getCreationHelper().createFormulaEvaluator();
			d = formulaEval.evaluate(cell).getNumberValue();
			break;
		case Cell.CELL_TYPE_NUMERIC:
			// 数字格式
			d = Double.parseDouble(NumberToTextConverter.toText(cell.getNumericCellValue()));
			break;
		case Cell.CELL_TYPE_STRING:
			d = Double.parseDouble(getString(cell));
			break;
		default:
			d = 0;
			break;
		}
		return d;

	}

	/**
	 * 获取字符串
	 * 
	 * @param cell
	 * @return
	 */
	@SuppressWarnings("deprecation")
	private static String getString(Cell cell) {
		if (cell == null) {
			return "";
		}
		if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
			return "";
		}
		cell.setCellType(Cell.CELL_TYPE_STRING);
		return cell.getStringCellValue().toString();
	}

	/**
	 * 对excel 2003处理
	 */
	private static Workbook xls(InputStream is) {
		try {
			// 得到工作簿
			return new HSSFWorkbook(is);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * 对excel 2007处理
	 */
	private static Workbook xlsx(InputStream is) {
		try {
			// 得到工作簿
			return new XSSFWorkbook(is);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * 创建单元格
	 * 
	 * @param row
	 * @param c
	 * @param cellValue
	 * @param style
	 */
	private static void creCell(Row row, int c, String cellValue, CellStyle style) {
		Cell cell = row.createCell(c);
		cell.setCellValue(cellValue);
		cell.setCellStyle(style);
	}

	public static void main(String[] args) {
		
		/**
		 * 输出文件
		 * */
//		ExcelEntity e = new ExcelEntity(5, "111");
//		ExcelEntity e1 = new ExcelEntity(2, "222");
//		List<ExcelEntity> list = new ArrayList<ExcelEntity>();
//		list.add(e);
//		list.add(e1);
//		String strPath = "/home/wenruo/Desktop/TestExcel/test2.xls";  	
//		File file = new File(strPath);  
//		try {
//			file.createNewFile();
//		} catch (IOException e2) {
//			// TODO Auto-generated catch block
//			e2.printStackTrace();
//		}
//		
//		outExcelFile(list, strPath);
		/**
		 * 读取文件 
		 * */
		List<ExcelEntity> excelEntities = readExcelFile("/home/wenruo/Desktop/TestExcel/student.xls", ExcelEntity.class);
		for (ExcelEntity excelEntity : excelEntities) {
			System.out.println(excelEntity.toString());
		}
	}

}
