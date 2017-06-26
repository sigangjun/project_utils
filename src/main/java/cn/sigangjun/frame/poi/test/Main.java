package cn.sigangjun.frame.poi.test;

import org.apache.log4j.Logger;

import java.io.IOException;
import java.util.List;

import cn.sigangjun.frame.poi.ExcelUtils;

public class Main {
	/**
	 * Logger for this class
	 */
	private static final Logger logger = Logger.getLogger(Main.class);

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
			logger.info(dto);
		}

	}

}
