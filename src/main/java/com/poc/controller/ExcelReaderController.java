package com.poc.controller;

import com.poc.service.ExcelService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@Slf4j
public class ExcelReaderController {

	@Autowired
	private ExcelService excelService;

	@PostMapping(value = "/savedata", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
	public ResponseEntity<Object> readExcelSheet (@RequestPart("file") MultipartFile file) {
		return ResponseEntity.ok().body(excelService.readXlsxFile(file));
	}

	@GetMapping("/writedata/{fileName}")
	public ResponseEntity<Object> writeDataInXlsx(@PathVariable String fileName){
		return ResponseEntity.ok(excelService.writeDataBaseDataInXlsx(fileName));
	}
}
