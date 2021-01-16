package com.apache.poi.model;

import com.fasterxml.jackson.annotation.JsonProperty;
import org.springframework.boot.context.properties.bind.Name;

import java.time.LocalDate;

@lombok.Data
public class Data {
    private Integer id = 1;
    @JsonProperty("NAME")
    private String name = "Nước";
    @JsonProperty("PRICE")
    private Integer price = 10_000;
    private boolean avail = true;
    private LocalDate date = LocalDate.of(2020,12,1);
}
