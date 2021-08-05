package com.poc.service;

import com.poc.entity.Family;
import com.poc.repo.FamilyRepository;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

@Service
@Slf4j
public class ExcelService {

	@Autowired
	private FamilyRepository familyRepository;

	public static void main (String[] args) {
		new ExcelService().writeDataBaseDataInXlsx("Family.xlsx");
	}

	public Object readXlsxFile (MultipartFile inputFile) {
		Family family = new Family();
		try (Workbook workbook = WorkbookFactory.create(inputFile.getInputStream())) {
			FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
			Sheet sheet = workbook.getSheetAt(0);
			for (Row row : sheet) {
				for (Cell cell : row) {
					switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
						case NUMERIC:
							System.out.print(Math.round(cell.getNumericCellValue()) + "\t\t");
							break;
						case STRING:
							System.out.print(cell.getStringCellValue() + "\t\t");
							break;
					}
				}
				System.out.println();
				if ( row.getRowNum() > 0 ) {
					family.setId((long) row.getCell(0).getNumericCellValue());
					family.setName(row.getCell(1).getStringCellValue());
					family.setRelation(row.getCell(2).getStringCellValue());
					familyRepository.save(family);
				}
			}
		} catch (IOException e) {
			log.info("Exception Occurred: e -> {}", e.getMessage(), e);
		}
		return familyRepository.findAll();
	}

	public Object writeDataBaseDataInXlsx (String fileName) {

		FileOutputStream out;
		try (Workbook workbook = new XSSFWorkbook()) {
			Sheet sheet = workbook.createSheet("Family");

			List<Family> familyList = familyRepository.findAll();
			Object[] headers = new Object[] {"Sl.No", "Name", "Relation"};

			Map<Long, Object[]> data = new TreeMap<>();
			data.put(1L, headers);

			for (Family family : familyList) {
				data.put(family.getId() + 1, new Object[] {family.getId(), family.getName(), family.getRelation()});
			}
			Set<Long> keyset = data.keySet();
			int rownum = 0;
			for (Long key : keyset) {
				Row row = sheet.createRow(rownum++);
				Object[] objArr = data.get(key);
				int cellnum = 0;
				for (Object obj : objArr) {
					Cell cell = row.createCell(cellnum++);
					if ( obj instanceof String )
						cell.setCellValue((String) obj);
					else if ( obj instanceof Long )
						cell.setCellValue((Long) obj);
				}
			}
			//Write the workbook in file system
			out = new FileOutputStream(new File("D://" + fileName));
			workbook.write(out);
			out.close();
			System.out.println("Family.xlsx written successfully on disk.");

		} catch (Exception e) {
			log.error("Exception Occurred : e -> {}", e.getMessage(), e);
		}

		return "File Location -> D://" + fileName;
	}
}