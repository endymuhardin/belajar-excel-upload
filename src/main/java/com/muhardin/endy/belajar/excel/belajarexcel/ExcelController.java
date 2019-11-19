package com.muhardin.endy.belajar.excel.belajarexcel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;

@Controller @Slf4j
public class ExcelController {

    @GetMapping("/upload/form")
    public void displayFormUpload() {

    }

    @PostMapping("/upload/form")
    public String prosesFormUpload(
            @RequestParam("nilai") MultipartFile fileNilai,
            RedirectAttributes redirectAttrs) {
        log.debug("Nama file : {}",fileNilai.getOriginalFilename());
        log.debug("Ukuran file : {} bytes", fileNilai.getSize());

        List<NilaiDto> hasilPenilaian = new ArrayList<>();

        try {
            Workbook workbook = new XSSFWorkbook(fileNilai.getInputStream());
            Sheet sheetPertama = workbook.getSheetAt(0);
            int rowPertama = 6;
            int jumlahMahasiswa = 10;
            for (int i = 0; i < jumlahMahasiswa; i++) {
                Row baris = sheetPertama.getRow(rowPertama + i);

                NilaiDto nilaiDto = new NilaiDto();

                Cell cellNim = baris.getCell(0); // tipe cell nim bisa string atau numeric, harus dihandle
                if (CellType.NUMERIC.equals(cellNim.getCellType())) {
                    nilaiDto.setNim(new BigDecimal(
                            cellNim.getNumericCellValue())
                            .setScale(0, RoundingMode.UNNECESSARY)
                            .toString());
                } else if (CellType.STRING.equals(cellNim.getCellType())) {
                    nilaiDto.setNim(cellNim.getStringCellValue());
                }

                nilaiDto.setNama(baris.getCell(1).getStringCellValue());
                nilaiDto.setNilai(new BigDecimal(baris.getCell(2).getNumericCellValue()));

                hasilPenilaian.add(nilaiDto);
            }
        } catch (IOException e) {
            log.error(e.getMessage(), e);
        }

        redirectAttrs.addFlashAttribute("hasilPenilaian", hasilPenilaian);
        return "redirect:hasil";
    }

    @GetMapping("/upload/hasil")
    public ModelMap tampilkanHasilUpload(@ModelAttribute("hasilPenilaian") List<NilaiDto> daftarNilai) {
        return new ModelMap()
                .addAttribute("daftarNilai", daftarNilai);
    }
}
