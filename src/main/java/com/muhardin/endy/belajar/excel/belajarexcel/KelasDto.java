package com.muhardin.endy.belajar.excel.belajarexcel;

import lombok.Data;

import javax.validation.constraints.Min;
import javax.validation.constraints.Size;

@Data
public class KelasDto {

    @Size(min = 3, max = 10)
    private String kodeMatakuliah;

    @Size(min = 3, max = 10)
    private String namaMatakuliah;

    @Size(min = 3, max = 10)
    private String namaDosen;

    @Min(5)
    private Integer jumlahMahasiswa = 20;
}
