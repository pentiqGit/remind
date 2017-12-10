package com.example.remind.controller;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


@SpringBootApplication
@RestController
public class RemindController {

	@RequestMapping(value="/")
	private String index() {
		String INPUT_DIR = "D:\\新しいフォルダー\\";
		String xlsxFileAddress = INPUT_DIR + "Book1.xlsx";

		List<List<String>>  rowList = new ArrayList<>();
		List<String> cellList = null;

		try (Workbook workbook = WorkbookFactory.create(new FileInputStream(xlsxFileAddress));) {
			Sheet sheet1 = workbook.getSheetAt(0);

			for (int i = 0; i < 10 ; i++) {
				Row row = sheet1.getRow(i);
				if (row == null) {
					continue;
				}

				cellList = new ArrayList<>();

				for (int j = 0; j < 10 ; j++) {
					Object cell = row.getCell(j);
					if (cell != null) {
						cell = getParse(row.getCell(j));
					}
						cellList.add(cell.toString());
				}
				rowList.add(cellList);
			}

			System.out.println("rowList.toString()=" + rowList.toString());

		} catch (Exception e) {
			e.printStackTrace();
		}

		return "";
	}

	private Object getParse(Cell cell) {
		switch (cell.getCellTypeEnum()) {
			case BLANK:
				return "";
			case _NONE:
				throw new RuntimeException("cell is null");
			case STRING:
				return cell.getStringCellValue();
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell))  {
					return cell.getDateCellValue(); // Date型
				} else {
					return cell.getNumericCellValue(); // double型
				}
			case BOOLEAN:
				return cell.getBooleanCellValue();
			case ERROR:
				throw new RuntimeException("Error cell is unsupported");
			case FORMULA:
				throw new RuntimeException("Formula cell is unsupported");
		}

		return null;
	}
}
