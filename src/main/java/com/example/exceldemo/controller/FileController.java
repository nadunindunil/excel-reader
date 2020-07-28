package com.example.exceldemo.controller;

import com.example.exceldemo.model.FormattingRun;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

// https://stackoverflow.com/questions/18760416/difference-between-hssfworkbook-and-workbook-in-apache-poi#:~:text=HSSFWorkbook%3A%20This%20class%20has%20methods,xls%20format.&text=XSSFWorkbook%3A%20This%20class%20has%20methods,and%20OpenOffice%20xml%20files%20in%20.

@RestController
public class FileController {
    @PostMapping("/upload")
    public void uploadFile(@RequestParam("file") MultipartFile mFile) throws IOException {
        // xls type files parser
        File file = new File(System.getProperty("java.io.tmpdir")+"/"+mFile.getOriginalFilename());
        mFile.transferTo(file);
        FileInputStream inputStream = new FileInputStream(file);

        Workbook wb = new HSSFWorkbook(inputStream);

        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);

        HSSFRichTextString richTextString = (HSSFRichTextString) cell.getRichStringCellValue();

        List<FormattingRun> formattingRuns = new ArrayList<FormattingRun>();
        int numFormattingRuns = richTextString.numFormattingRuns();
        for (int fmtIdx = 0; fmtIdx < numFormattingRuns;fmtIdx++)
        {
            int begin = richTextString.getIndexOfFormattingRun(fmtIdx);
            short fontIndex = richTextString.getFontOfFormattingRun(fmtIdx);

            // Walk the string to determine the length of the formatting run.
            int length = 0;
            for (int j = begin; j < richTextString.length(); j++)
            {
                short currFontIndex = richTextString.getFontAtIndex(j);
                if (currFontIndex == fontIndex)
                    length++;
                else
                    break;
            }
            formattingRuns.add(new FormattingRun(begin, length, fontIndex));
        }

        System.out.println(richTextString);
        System.out.println(formattingRuns);
    }

    @GetMapping("download")
    public ResponseEntity<Resource> downloadFile() throws IOException {
        //  Resources
        //  https://stackoverflow.com/questions/22093864/java-net-malformedurlexception-no-protocol-on-url-based-on-a-string-modified-wi
        //  https://stackoverflow.com/questions/26860167/what-is-a-safe-way-to-create-a-temp-file-in-java

        String contentType = "application/vnd.ms-excel";
        String fileName = "testFile";

        File tempFile = File.createTempFile("prefix-", ".xls");
        tempFile.deleteOnExit();
        Workbook wb = new HSSFWorkbook();

        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);

        Font font = wb.createFont();
        font.setBold(true);

        RichTextString richString = new HSSFRichTextString( "This is a test" );
        //^0 ^3 ^6 ^9
        richString.applyFont(0,5, font);
        cell.setCellValue( richString );

        FileOutputStream fileOut = new FileOutputStream(tempFile);
        wb.write(fileOut);
        wb.close();

        Path filePath = Paths.get(tempFile.getPath()).toAbsolutePath().normalize();

        Resource resource = new UrlResource(filePath.toUri());

        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType(contentType))
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + fileName  + ".xls\"")
                .body(resource);
    }
}
