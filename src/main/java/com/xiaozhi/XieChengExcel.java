package com.xiaozhi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class XieChengExcel {

    static List<String> CITY_LIST = Arrays.asList("西安", "兰州", "延安", "西宁", "汉中", "天水", "喀什市", "乌鲁木齐", "安康", "宝鸡", "固原", "平凉", "吐鲁番市", "榆林", "银川");

    static SimpleDateFormat FORMATTER = new SimpleDateFormat("yyyy-MM");

    public static void main(String[] args) throws Exception {
        List<XieChengData> xieChengDataList = readExcel();
        List<CityHotelStatistics> statisticsList = handleData(xieChengDataList);
        writeExcel(statisticsList);
    }

    /**
     * 写excel
     */
    private static void writeExcel(List<CityHotelStatistics> statisticsList) throws Exception {
        Set<String> brandSet = statisticsList.stream().map(k -> k.brand).collect(Collectors.toSet());
        List<String> brandList = new ArrayList<>(brandSet);
        XSSFWorkbook workbook = new XSSFWorkbook();
        File file = new File("/Users/huangzhi/PythonProject/竞品酒店/竞品酒店西北月度数据.xlsx");

        Sheet sheet = workbook.createSheet();
        //首行
        Row titleRow = sheet.createRow(0);
        Cell cell0 = titleRow.createCell(0);
        cell0.setCellValue("城市");
        Cell cell1 = titleRow.createCell(1);
        cell1.setCellValue("月份");
        for (int i = 0; i < brandList.size(); i++) {
            Cell cellI = titleRow.createCell(i + 2);
            cellI.setCellValue(brandList.get(i));
        }
        //数据
        Map<String, List<CityHotelStatistics>> cityMonthMapList = statisticsList.stream().collect(Collectors.groupingBy(a -> a.city + "丢" + a.month));
        AtomicInteger excelRowIndex = new AtomicInteger(1);
        cityMonthMapList.forEach((k, v) -> {
            String[] split = k.split("丢");
            String city = split[0];
            String month = split[1];
            Row row = sheet.createRow(excelRowIndex.get());
            Cell cellData0 = row.createCell(0);
            cellData0.setCellValue(city);
            Cell cellData1 = row.createCell(1);
            cellData1.setCellValue(month);
            for (CityHotelStatistics statistics : v) {
                int i = brandList.indexOf(statistics.brand);
                Cell cellI = row.createCell(i + 2);
                cellI.setCellValue(statistics.score);
            }
            excelRowIndex.getAndIncrement();
        });

        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
    }

    /**
     * 转化成月份
     */
    public static String getMonth(Date date) {
        return FORMATTER.format(date);
    }

    /**
     * 计算月平均分
     */
    private static String getScore(List<XieChengData> monthList) {
        double sum = 0;
        double count = 0;
        for (XieChengData xieChengData : monthList) {
            sum += xieChengData.score * xieChengData.count;
            count += xieChengData.count;
        }
        String result = "0";
        try {
            result = new BigDecimal(sum).divide(new BigDecimal(count),2,RoundingMode.HALF_UP).toString();
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("-----------------");
            System.out.println(monthList.get(0));
        }
        return result;
    }

    /**
     * 处理解析数据
     */
    public static List<CityHotelStatistics> handleData(List<XieChengData> xieChengDataList) {
        List<CityHotelStatistics> statisticsList = new ArrayList<>();
        Map<String, List<XieChengData>> cityMapList = xieChengDataList.stream().collect(Collectors.groupingBy(a -> a.getCity()));
        cityMapList.forEach((city, cityList) -> {
            Map<String, List<XieChengData>> brandDataList = cityList.stream().collect(Collectors.groupingBy(a -> a.getBrand()));
            brandDataList.forEach((brand, brandList) -> {
                Map<String, List<XieChengData>> monthMapList = brandList.stream().collect(Collectors.groupingBy(a -> getMonth(a.getDateTime())));
                monthMapList.forEach((month, monthList) -> {
                    CityHotelStatistics statistics = new CityHotelStatistics();
                    statistics.city = city;
                    statistics.brand = brand;
                    statistics.month = month;
                    statistics.score = getScore(monthList);
                    statisticsList.add(statistics);
                });
            });
        });
        return statisticsList;
    }

    public static List<XieChengData> readExcel() throws Exception {
        List<XieChengData> xieChengDataList = new ArrayList<>();
        File file = new File("/Users/huangzhi/PythonProject/竞品酒店/竞品酒店-携程.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = workbook.getSheetAt(0);
        System.out.println("总行数：" + sheet.getLastRowNum());
        System.out.println("物理行数" + sheet.getPhysicalNumberOfRows());

        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            XSSFRow row = sheet.getRow(i);
            String city = row.getCell(4).getStringCellValue();
            if (!CITY_LIST.contains(city)) {
                continue;
            }
            XieChengData xieChengData = new XieChengData();
            xieChengData.hotelName = row.getCell(1).getStringCellValue();
            xieChengData.dateTime = row.getCell(2).getDateCellValue();
            xieChengData.brand = row.getCell(3).getStringCellValue();
            xieChengData.city = city;
            xieChengData.score = row.getCell(6).getNumericCellValue();
            xieChengData.count = row.getCell(7).getNumericCellValue();
            xieChengDataList.add(xieChengData);
        }
        workbook.close();
        return xieChengDataList;
    }
}
