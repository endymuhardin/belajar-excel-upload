package com.muhardin.endy.belajar.excel.belajarexcel;

import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

@Controller @Slf4j
public class ExcelController {

    @GetMapping("/upload/form")
    public void displayFormUpload() {

    }

    @PostMapping("/upload/form")
    public String prosesFormUpload(@RequestParam("nilai") MultipartFile fileNilai) {
        log.debug("Nama file : {}",fileNilai.getOriginalFilename());
        log.debug("Ukuran file : {} bytes", fileNilai.getSize());

        return "redirect:hasil";
    }

    @GetMapping("/upload/hasil")
    public ModelMap tampilkanHasilUpload() {
        List<NilaiDto> daftarNilai = new ArrayList<>();

        NilaiDto n = new NilaiDto();
        n.setNim("123");
        n.setNama("Test 123");
        n.setNilai(new BigDecimal("75.5"));
        daftarNilai.add(n);
        return new ModelMap()
                .addAttribute("daftarNilai", daftarNilai);
    }
}
