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

import javax.servlet.http.HttpServletResponse;
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

    @GetMapping("/template-nilai.xlsx")
    public void downloadTemplateNilai(HttpServletResponse response) throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Nilai UTS");

        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 15000);

        createHeader(sheet);

        List<MahasiswaDto> daftarMahasiswa = generateDaftarMahasiswa();
        createDaftarNilai(sheet, daftarMahasiswa);

        response.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment;filename=nilai-uts.xlsx");
        wb.write(response.getOutputStream());
    }

    private List<MahasiswaDto> generateDaftarMahasiswa() {
        List<MahasiswaDto> daftarMahasiswa = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            daftarMahasiswa.add(
                    new MahasiswaDto("1234567890" + i, "Mahasiswa "+i));
        }

        return  daftarMahasiswa;
    }

    private void createHeader(Sheet sheet) {
        CellStyle style = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontName("arial");
        style.setFont(font);

        Row row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("Kode Matakuliah");
        row1.createCell(1).setCellValue("FM-001");

        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("Nama Matakuliah");
        row2.createCell(1).setCellValue("Fiqh Muamalah");

        Row row3 = sheet.createRow(2);
        row3.createCell(0).setCellValue("Nama Dosen");
        row3.createCell(1).setCellValue("Abdul Mughni, Lc. MHi");

        row1.getCell(0).setCellStyle(style);
        row1.getCell(1).setCellStyle(style);
        row2.getCell(0).setCellStyle(style);
        row2.getCell(1).setCellStyle(style);
        row3.getCell(0).setCellStyle(style);
        row3.getCell(1).setCellStyle(style);

    }

    private void createDaftarNilai(Sheet sheet, List<MahasiswaDto> daftarMahasiswa) {
        CellStyle styleContent = sheet.getWorkbook().createCellStyle();
        styleContent.setBorderBottom(BorderStyle.THIN);
        styleContent.setBorderLeft(BorderStyle.THIN);
        styleContent.setBorderRight(BorderStyle.THIN);
        styleContent.setBorderTop(BorderStyle.THIN);

        CellStyle styleHeader = sheet.getWorkbook().createCellStyle();
        styleHeader.cloneStyleFrom(styleContent);

        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        styleHeader.setFont(font);

        // header
        Row hdr = sheet.createRow(5);
        hdr.createCell(0).setCellValue("NIM");
        hdr.createCell(1).setCellValue("Nama Mahasiswa");
        hdr.createCell(2).setCellValue("Nilai");

        hdr.getCell(0).setCellStyle(styleHeader);
        hdr.getCell(1).setCellStyle(styleHeader);
        hdr.getCell(2).setCellStyle(styleHeader);

        // data mahasiswa
        int baris = 5;
        for (MahasiswaDto m : daftarMahasiswa) {
            baris++;

            Row mhs = sheet.createRow(baris);
            mhs.createCell(0).setCellValue(m.getNim());
            mhs.createCell(1).setCellValue(m.getNama());

            mhs.getCell(0).setCellStyle(styleContent);
            mhs.getCell(1).setCellStyle(styleContent);
            mhs.createCell(2).setCellStyle(styleContent);
        }
    }

    @PostMapping("/upload/form")
    public String prosesFormUpload(
            @RequestParam("nilai") MultipartFile fileNilai,
            RedirectAttributes redirectAttrs) {

        log.debug("Nama file : {}", fileNilai.getOriginalFilename());
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
