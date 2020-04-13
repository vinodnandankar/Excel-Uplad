package com.example.demo.controller;

import java.util.List;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.model.DataSheet;
import com.example.demo.service.UploadSheetService;

@RestController
public class UploadController {
	
	public final UploadSheetService uploadSheetService;

	public UploadController(UploadSheetService uploadSheetService) {
		this.uploadSheetService=uploadSheetService;
	}

	
	@PostMapping("/upload")
	public List<DataSheet> upload(@RequestParam("file")MultipartFile file) throws Exception {
		return uploadSheetService.upload(file);

	}
}
