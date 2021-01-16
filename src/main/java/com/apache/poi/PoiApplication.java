package com.apache.poi;

import com.apache.poi.model.Data;
import com.apache.poi.service.PoiService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

//@SpringBootApplication
public class PoiApplication {

    public static void main(String[] args) throws IOException, IllegalAccessException {
//        SpringApplication.run(PoiApplication.class, args);
        String readFile = "src/main/resources/static/file/readData.xlsx";
        String writeFile = "src/main/resources/static/file/writeData.xlsx";
        PoiService poiService = new PoiService();
        poiService.readData(readFile);
        List<Data> dataList = new ArrayList<>();
        dataList.add(new Data());
        poiService.writeData(writeFile, dataList);
    }

}
