package com.uray;

import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelTest {

    private static void treateExcelData() {
        // 读Excel表，取得数据并处理
        String sourceName = "E:\\symdata_work\\jd_data\\sale\\20190212_all.xlsx";
        List<List<Object>> dataList = readExcel(sourceName);
        System.out.println("------read executed------");
        // 将数据写入新的Excel表
        String fileName = "E:\\symdata_work\\jd_data\\result\\京东商城交易数据20190212.xlsx";
        writeExcel(fileName, dataList);
        System.out.println("------write executed------");
    }

    private static void writeExcel(String fileName, List<List<Object>> content) {
        String[] colNames = { "日期", "时间", "下单数", "成功订单数", "下单-成功转化率", "下单用户数", "成功用户数", "下单用户转化率", "全部下单金额", "成功订单金额",
                "备注" };
        String[] CPColNames = { "使用电子券金额", "成功订单数", "成功付款金额", "客单价" };
        String[] CPRowNames = { "0元", "0.01-9.99元", "10-19.99元", "20-29.99元", "30-39.99元", "40-49.99元", "50-99.99元",
                "100元以上" };
        String[] transFreeRowNames = { "49元以下", "49-98.99元", "99元及以上" };
        String[] transFreeColNames = { "付款金额", "订单数量", "订单金额", "客单价" };

        Workbook saleWorkbook = null;
        if (fileName.endsWith(".xlsx")) {
            saleWorkbook = new XSSFWorkbook();
        } else if (fileName.endsWith(".xls")) {
            saleWorkbook = new HSSFWorkbook();
        }

        // 写入销售数据
        Sheet saleSheet = saleWorkbook.createSheet("jdsale");
        Row saleRow = saleSheet.createRow(0);
        Cell saleCell = null;
        for (int colIndex = 0; colIndex < colNames.length; colIndex++) {// 写表头
            saleCell = saleRow.createCell(colIndex);
            saleCell.setCellValue(colNames[colIndex]);
        }
        List<Object> saleDataList = content.get(0);
        Row saleDataRow = saleSheet.createRow(1);
        Cell saleDataCell = null;
        for (int colIndex = 0; colIndex < colNames.length; colIndex++) {
            saleDataCell = saleDataRow.createCell(colIndex);
            if (colIndex == 4 || colIndex == 7) {// 百分数格式
                saleDataCell.setCellValue(Double.parseDouble(String.valueOf(saleDataList.get(colIndex))));
                CellStyle cellStyle = saleWorkbook.createCellStyle();
                cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00%"));
                saleDataCell.setCellStyle(cellStyle);
            } else if (colIndex == 8 || colIndex == 9) {// 保留两位小数
                saleDataCell.setCellValue(Double.parseDouble(String.valueOf(saleDataList.get(colIndex))));
                CellStyle cellStyle = saleWorkbook.createCellStyle();
                cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                saleDataCell.setCellStyle(cellStyle);
            } else if (colIndex == 0 || colIndex == 2 || colIndex == 3 || colIndex == 5 || colIndex == 6) {
                saleDataCell.setCellValue(Integer.parseInt(String.valueOf(saleDataList.get(colIndex))));
            } else {
                saleDataCell.setCellValue(String.valueOf(saleDataList.get(colIndex)));
            }
        }

        // 写入电子券数据
        Sheet couponSheet = saleWorkbook.createSheet("jdcoupon");
        Row couponRow = couponSheet.createRow(0);
        Cell couponCell = null;
        for (int colIndex = 0; colIndex < CPColNames.length; colIndex++) {// 写表头
            couponCell = couponRow.createCell(colIndex);
            couponCell.setCellValue(CPColNames[colIndex]);
        }
        List<Object> couponDataList = content.get(1);
        Row couponDataRow = null;
        Cell couponDataCell = null;
        int i = 0;
        for (int rowIndex = 1; rowIndex <= CPRowNames.length; rowIndex++) {
            couponDataRow = couponSheet.createRow(rowIndex);
            for (int colIndex = 0; colIndex < CPColNames.length; colIndex++) {
                couponDataCell = couponDataRow.createCell(colIndex);
                if (colIndex == 0) {// 插入行名
                    couponDataCell.setCellValue(CPRowNames[i]);
                    i++;
                } else if (colIndex == 1) {
                    couponDataCell.setCellValue((int) couponDataList.get(2 * rowIndex - 2));
                } else if (colIndex == 2) {
                    couponDataCell.setCellValue((double) couponDataList.get(2 * rowIndex - 1));
                    CellStyle cellStyle = saleWorkbook.createCellStyle();
                    cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                    couponDataCell.setCellStyle(cellStyle);
                } else if (colIndex == 3) {
                    String formula = "C" + (rowIndex + 1) + "/" + "B" + (rowIndex + 1);
                    couponDataCell.setCellFormula(formula);
                    couponDataCell.setCellType(Cell.CELL_TYPE_FORMULA);
                    CellStyle cellStyle = saleWorkbook.createCellStyle();
                    cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                    couponDataCell.setCellStyle(cellStyle);
                }
            }
        }

        // 写入99元包邮数据
        Sheet transFreeSheet = saleWorkbook.createSheet("transFree");
        Row transFreeRow = transFreeSheet.createRow(0);
        Cell transFreeCell = null;
        for (int colIndex = 0; colIndex < transFreeColNames.length; colIndex++) {// 写表头
            transFreeCell = transFreeRow.createCell(colIndex);
            transFreeCell.setCellValue(transFreeColNames[colIndex]);
        }
        List<Object> transFreeDataList = content.get(2);
        Row transFreeDataRow = null;
        Cell transFreeDataCell = null;
        int j = 0;
        for (int rowIndex = 1; rowIndex <= transFreeRowNames.length; rowIndex++) {
            transFreeDataRow = transFreeSheet.createRow(rowIndex);
            for (int colIndex = 0; colIndex < transFreeColNames.length; colIndex++) {
                transFreeDataCell = transFreeDataRow.createCell(colIndex);
                if (colIndex == 0) {// 插入行名
                    transFreeDataCell.setCellValue(transFreeRowNames[j]);
                    j++;
                } else if (colIndex == 1) {
                    transFreeDataCell.setCellValue((int) transFreeDataList.get(2 * rowIndex - 2));
                } else if (colIndex == 2) {
                    transFreeDataCell.setCellValue((double) transFreeDataList.get(2 * rowIndex - 1));
                    CellStyle cellStyle = saleWorkbook.createCellStyle();
                    cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                    transFreeDataCell.setCellStyle(cellStyle);
                } else if (colIndex == 3) {
                    String formula = "C" + (rowIndex + 1) + "/" + "B" + (rowIndex + 1);
                    transFreeDataCell.setCellFormula(formula);
                    transFreeDataCell.setCellType(Cell.CELL_TYPE_FORMULA);
                    CellStyle cellStyle = saleWorkbook.createCellStyle();
                    cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                    transFreeDataCell.setCellStyle(cellStyle);
                }
            }
        }

        // saleSheet.setForceFormulaRecalculation(true);// 格式自动生效

        // 写文件
        File saleFile = new File(fileName);
        try {
            saleFile.createNewFile();
            FileOutputStream fos = new FileOutputStream(saleFile);
            saleWorkbook.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static List<List<Object>> readExcel(String sourceName) {
        String date = "";
        String day = "";
        int ordNum = 0;
        int succOrdNum = 0;
        int userNum = 0;
        int succUserNum = 0;
        double totalFee = 0;
        double payedFee = 0;
        double ordPrice = 0;
        // double ordPriForCP = 0;
        String mark = "";
        List<Double> ids = new ArrayList<>();
        boolean userFlag = false;// 非重复用户标记
        boolean payedFlag = false;// 订单支付标记

        // 订单-电子券
        int ordNumNoCP = 0;
        int ordNumCPIn10 = 0;
        int ordNumCPIn10_20 = 0;
        int ordNumCPIn20_30 = 0;
        int ordNumCPIn30_40 = 0;
        int ordNumCPIn40_50 = 0;
        int ordNumCPIn50_100 = 0;
        int ordNumCPOver100 = 0;
        double feeNoCP = 0;
        double feeCPIn10 = 0;
        double feeCPIn10_20 = 0;
        double feeCPIn20_30 = 0;
        double feeCPIn30_40 = 0;
        double feeCPIn40_50 = 0;
        double feeCPIn50_100 = 0;
        double feeCPInOver100 = 0;
        boolean payedUCoupon = false;
        // 包邮表格数据
        int ordNum99In49 = 0;
        double fee99In49 = 0;
        int ordNum99In49_99 = 0;
        double fee99In49_99 = 0;
        int ordNum99Over99 = 0;
        double fee99Over99 = 0;
        Workbook workbook = null;
        try {
            InputStream is = new FileInputStream(sourceName);
            if (sourceName.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(is);
            } else if (sourceName.endsWith(".xls")) {
                workbook = new HSSFWorkbook(is);
            }
            Sheet sheet = workbook.getSheetAt(0);
            ordNum = sheet.getLastRowNum();// 订单数量
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }
                for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    switch (cellIndex) {
                        case 0:// 用户ID
                            if (cell.getNumericCellValue() != 0 && !ids.contains(cell.getNumericCellValue())) {
                                ids.add(cell.getNumericCellValue());
                                userNum++;
                                userFlag = true;
                            } else {
                                userFlag = false;
                            }

                            break;
                        case 1:// 日期
                            if (cell.getDateCellValue() != null && cell.getRow().getRowNum() == sheet.getLastRowNum()) {// 最后一行
                                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
                                date = dateFormat.format(cell.getDateCellValue());
                                SimpleDateFormat dateFm = new SimpleDateFormat("EEEE");
                                day = dateFm.format(cell.getDateCellValue());
                            }
                            break;
                        case 2:// 订单状态

                            break;
                        case 3:// 支付状态
                            if (cell != null && cell.getNumericCellValue() == 1) {// 判断成功订单
                                succOrdNum++;// 成功支付订单
                                payedFlag = true;// 成功支付标记
                                payedUCoupon = true;
                                if (userFlag) {// 判断未重复用户
                                    succUserNum++;
                                    userFlag = false;
                                }
                            } else {
                                payedFlag = false;
                                payedUCoupon = false;
                            }
                            break;
                        case 4:// 运费

                            break;
                        case 5:// 订单金额
                            ordPrice = cell.getNumericCellValue() / 100;
                            totalFee += ordPrice;// 全部下单金额
                            if (payedFlag) {// 成功订单金额
                                payedFee += ordPrice;
                                if (ordPrice < 49) {
                                    ordNum99In49++;
                                    fee99In49 += ordPrice;
                                } else if (ordPrice >= 49 && ordPrice < 99) {
                                    ordNum99In49_99++;
                                    fee99In49_99 += ordPrice;
                                } else {
                                    ordNum99Over99++;
                                    fee99Over99 += ordPrice;
                                }
                                payedFlag = false;// 支付标记置为false
                            }
                            break;
                        case 6:// 电子券
                            if (payedUCoupon) {
                                double coupon = cell.getNumericCellValue() / 100;
                                if (coupon <= 0) {
                                    ordNumNoCP++;
                                    feeNoCP += ordPrice;
                                } else if (coupon < 10) {
                                    ordNumCPIn10++;
                                    feeCPIn10 += ordPrice;
                                } else if (coupon < 20) {
                                    ordNumCPIn10_20++;
                                    feeCPIn10_20 += ordPrice;
                                } else if (coupon < 30) {
                                    ordNumCPIn20_30++;
                                    feeCPIn20_30 += ordPrice;
                                } else if (coupon < 40) {
                                    ordNumCPIn30_40++;
                                    feeCPIn30_40 += ordPrice;
                                } else if (coupon < 50) {
                                    ordNumCPIn40_50++;
                                    feeCPIn40_50 += ordPrice;
                                } else if (coupon < 100) {
                                    ordNumCPIn50_100++;
                                    feeCPIn50_100 += ordPrice;
                                } else {
                                    ordNumCPOver100++;
                                    feeCPInOver100 += ordPrice;
                                }
                                payedUCoupon = false;
                            }
                            break;

                        default:
                            break;
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        double ordConvertRate = (double) succOrdNum / ordNum;
        double userConvertRate = (double) succUserNum / userNum;

        List<List<Object>> resultList = new ArrayList<>();

        // 销售数据汇总
        List<Object> saleList = new ArrayList<>();
        saleList.add(date);
        saleList.add(day);
        saleList.add(ordNum);
        saleList.add(succOrdNum);
        saleList.add(ordConvertRate);
        saleList.add(userNum);
        saleList.add(succUserNum);
        saleList.add(userConvertRate);
        saleList.add(totalFee);
        saleList.add(payedFee);
        saleList.add(mark);

        // 电子券统计
        List<Object> couponSaleList = new ArrayList<>();
        couponSaleList.add(ordNumNoCP);
        couponSaleList.add(feeNoCP);
        couponSaleList.add(ordNumCPIn10);
        couponSaleList.add(feeCPIn10);
        couponSaleList.add(ordNumCPIn10_20);
        couponSaleList.add(feeCPIn10_20);
        couponSaleList.add(ordNumCPIn20_30);
        couponSaleList.add(feeCPIn20_30);
        couponSaleList.add(ordNumCPIn30_40);
        couponSaleList.add(feeCPIn30_40);
        couponSaleList.add(ordNumCPIn40_50);
        couponSaleList.add(feeCPIn40_50);
        couponSaleList.add(ordNumCPIn50_100);
        couponSaleList.add(feeCPIn50_100);
        couponSaleList.add(ordNumCPOver100);
        couponSaleList.add(feeCPInOver100);

        // 99元包邮统计
        List<Object> transFreeList = new ArrayList<>();
        transFreeList.add(ordNum99In49);
        transFreeList.add(fee99In49);
        transFreeList.add(ordNum99In49_99);
        transFreeList.add(fee99In49_99);
        transFreeList.add(ordNum99Over99);
        transFreeList.add(fee99Over99);

        resultList.add(0, saleList);
        resultList.add(1, couponSaleList);
        resultList.add(2, transFreeList);
        return resultList;
    }

    static class WeeklySale implements Runnable {

        public WeeklySale() {
        }

        @Override
        public void run() {
            String sourceName = "E:\\symdata_work\\jd_data\\sale_weekly\\201811_18-24.xlsx";
            List<List<List<Object>>> data = readWeeklyExcel(sourceName);
            // 将数据写入新的Excel表
            String fileName = "E:\\symdata_work\\jd_data\\result_weekly\\京东商城周数据统计201811_18-24.xlsx";
            writeWeeklyExcel(fileName, data);
        }

        private void writeWeeklyExcel(String fileName, List<List<List<Object>>> data) {
            List<List<Object>> topProvinces = data.get(0);
            List<List<Object>> topSkus = data.get(1);
            List<List<Object>> topPrices = data.get(2);
            List<List<Object>> topCoupons = data.get(3);

            Workbook workbook = null;
            if (fileName.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook();
            } else if (fileName.endsWith(".xls")) {
                workbook = new HSSFWorkbook();
            }

            // 写入top五省
            String provSheetName = "省份top5";
            String provTitle = "订单数量top5省";
            String[] provColNames = {"序号", "省份名称", "订单数量"};
            Sheet topProvSheet = workbook.createSheet(provSheetName);
            Row topProvRow = null;
            Cell topProvCell = null;
            CellStyle cellStyle = workbook.createCellStyle();

            uniFormat(provSheetName, provTitle, provColNames, workbook, topProvSheet, topProvRow, topProvCell, cellStyle, 0, provColNames.length - 1);

            for (int rowIndex = 2; rowIndex < topProvinces.size() + 2; rowIndex++) {
                topProvRow = topProvSheet.createRow(rowIndex);
                for (int colIndex = 0; colIndex < provColNames.length; colIndex++) {
                    topProvCell = topProvRow.createCell(colIndex);
                    if (colIndex == 0 || colIndex == 2) {
                        topProvCell.setCellValue((Integer) topProvinces.get(rowIndex - 2).get(colIndex));
                        topProvCell.setCellStyle(cellStyle);
                    } else {
                        topProvCell.setCellValue(String.valueOf(topProvinces.get(rowIndex - 2).get(colIndex)));
                        topProvCell.setCellStyle(cellStyle);
                    }
                }
            }

            //写入topSKU
            String skuSheetName = "商品top10";
            String skuTitle = "热销商品top10";
            String[] skuColNames = {"SKU", "销量", "商品名称"};
            Sheet topSkuSheet = workbook.createSheet(skuSheetName);
            Row topSkuRow = null;
            Cell topSkuCell = null;

            uniFormat(skuSheetName, skuTitle, skuColNames, workbook, topSkuSheet, topSkuRow, topSkuCell, cellStyle, 0, skuColNames.length - 1);

            for (int rowIndex = 2; rowIndex < topSkus.size() + 2; rowIndex++) {
                topSkuRow = topSkuSheet.createRow(rowIndex);
                for (int colIndex = 0; colIndex < skuColNames.length; colIndex++) {
                    topSkuCell = topSkuRow.createCell(colIndex);
                    if (colIndex == 0 || colIndex == 1) {
                        topSkuCell.setCellValue((Long) topSkus.get(rowIndex - 2).get(colIndex));
                        topSkuCell.setCellStyle(cellStyle);
                    } else {
                        topSkuCell.setCellValue(String.valueOf(topSkus.get(rowIndex - 2).get(colIndex)));
                        topSkuCell.setCellStyle(cellStyle);
                    }
                }
            }

            //写入订单价格最高top10
            String topPricSheetName = "订单top10";
            String topPricitle = "订单价格排名top10";
            String[] topPricColNames = {"京东订单号", "订单日期", "运费", "订单金额", "电子券", "省份", "sku", "商品内容"};

            writeTopOrders(workbook, cellStyle, topPrices, topPricSheetName, topPricitle, topPricColNames);

            //写入电子券价格最高top10
            String topCoupSheetName = "用券top10";
            String topCoupitle = "用券排名top10";
            String[] topCoupColNames = {"京东订单号", "订单日期", "运费", "订单金额", "电子券", "省份", "sku", "商品内容"};

            writeTopOrders(workbook, cellStyle, topCoupons, topCoupSheetName, topCoupitle, topCoupColNames);

            // 写文件
            File saleFile = new File(fileName);
            try {
                saleFile.createNewFile();
                FileOutputStream fos = new FileOutputStream(saleFile);
                workbook.write(fos);
                fos.close();
                System.out.println("=========================== 完成写表操作 ============================");
            } catch (IOException e) {
                e.printStackTrace();
            }

        }

        /**
         * @Description: 通用写订单
         * @author: Uray
         * @date: 5:56:52 PM Nov 23, 2018
         */
        private void writeTopOrders(Workbook workbook, CellStyle cellStyle, List<List<Object>> data, String sheetName, String title, String[] colNames) {
            Sheet sheet = workbook.createSheet(sheetName);
            sheet.autoSizeColumn(0);
            Row row = null;
            Cell cell = null;

            uniFormat(sheetName, title, colNames, workbook, sheet, row, cell, cellStyle, 0, colNames.length - 1);

            for (int rowIndex = 2; rowIndex < data.size() + 2; rowIndex++) {
                row = sheet.createRow(rowIndex);
                for (int colIndex = 0; colIndex < colNames.length; colIndex++) {
                    cell = row.createCell(colIndex);
                    if (colIndex == 5 || colIndex == 7) {
                        cell.setCellValue(String.valueOf(data.get(rowIndex - 2).get(colIndex)));
                        cell.setCellStyle(cellStyle);
                    } else if (colIndex == 0 || colIndex == 6) {
                        cell.setCellValue((Long) data.get(rowIndex - 2).get(colIndex));
                        cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0"));
                        cell.setCellStyle(cellStyle);
                    } else {
                        cell.setCellValue((Integer) data.get(rowIndex - 2).get(colIndex));
                        cell.setCellStyle(cellStyle);
                    }
                }
            }

        }

        /**
         * @Description: 约定格式
         * @author: Uray
         * @date: 2:36:50 PM Nov 22, 2018
         */
        private void uniFormat(String sheetName, String title, String[] colNames, Workbook workbook, Sheet sheet,
                               Row row, Cell cell, CellStyle cellStyle, int startCol, int endCol) {
            // 单元格格式
            cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
            cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            // 合并标题行
            // 参数1：行号，参数2：起始列号，参数3：行号，参数4：终止列号
            CellRangeAddress region = new CellRangeAddress(0, startCol, 0, endCol);
            sheet.addMergedRegion(region);
            // 写标题
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue(title);// 写标题
            cell.setCellStyle(cellStyle);
            // 写表头
            row = sheet.createRow(1);
            for (int colIndex = 0; colIndex < colNames.length; colIndex++) {// 写表头
                cell = row.createCell(colIndex);
                cell.setCellValue(colNames[colIndex]);
                cell.setCellStyle(cellStyle);
            }
        }

        /**
         * @Description: 读
         * @author: Uray
         * @date: 10:54:03 AM Nov 23, 2018
         */
        private List<List<List<Object>>> readWeeklyExcel(String sourceName) {
            List<Integer> provinceIdList = new ArrayList<>();
            for (int i = 0; i < 50; i++) {// 初始化省份ID集合
                provinceIdList.add(0);
            }

            Long sku = null;
            String name = null;
            Integer skuNum = 0;
            Long ordId = null;
            Integer orderDate = null;
            Integer freight = null;

            Map<Long, Integer> skuMap = new LinkedHashMap<>();
            Map<Long, String> skuNameMap = new LinkedHashMap<>();

            List<JDSearchOrderVo> topPriceOrders = new ArrayList<>();
            List<JDSearchOrderVo> topCouponOrders = new ArrayList<>();

            JDSearchOrderVo jdSearchOrder = new JDSearchOrderVo();
            JDSearchOrderVo maxPriceOrder = new JDSearchOrderVo();
            JDSearchOrderVo maxCouponOrder = new JDSearchOrderVo();

            List<Long> topPriceOrderIds = new ArrayList<>();
            List<Long> topCouponOrderIds = new ArrayList<>();
            Integer price = 0;
            Integer coupon = 0;
            boolean maxPayPrice = false;
            boolean maxCoupon = false;

            InputStream is = null;
            List<Long> jdOrdIdList = new ArrayList<>();// 订单号集合
            Map<Integer, String> provinceMap = new LinkedHashMap<>();// 省份id-名称对应关系
            boolean realOrder = false;//是否计入新的省份开关
            Integer provinceId = null;
            String provinceName = null;
            Integer provinceNum = null;

            Workbook workbook = null;
            Sheet sheet = null;
            Row row = null;

            try {
                is = new FileInputStream(sourceName);
                if (sourceName.endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(is);
                } else if (sourceName.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(is);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            sheet = workbook.getSheetAt(0);
            if (sheet != null) {
                for (int rowIndex = 1; rowIndex < sheet.getLastRowNum(); rowIndex++) {
                    row = sheet.getRow(rowIndex);
                    if (row == null) {
                        continue;
                    }
                    jdSearchOrder = new JDSearchOrderVo();
                    for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
                        if (colIndex == 1) {// 记录订单编号
                            ordId = (long) row.getCell(colIndex).getNumericCellValue();
                            if (!jdOrdIdList.contains(ordId)) {// 非重复单号
                                jdOrdIdList.add(ordId);
                                realOrder = true;
                            }
                        }
                        if (colIndex == 2) {
                            orderDate = (int) row.getCell(colIndex).getNumericCellValue();
                        }
                        if (colIndex == 5) {
                            freight = (int) row.getCell(colIndex).getNumericCellValue();
                        }
                        if (colIndex == 6) {// 订单价格
                            price = (int) row.getCell(colIndex).getNumericCellValue();
                            if (price > maxPriceOrder.getPayPrice()) {
                                maxPayPrice = true;
                            }
                        }
                        if (colIndex == 7) {// 筛选用券多的订单
                            coupon = (int) row.getCell(colIndex).getNumericCellValue();
                            if (coupon > maxCouponOrder.getCompanyPaymony()) {
                                maxCoupon = true;
                            }
                        }
                        if (colIndex == 8 && realOrder) {//处理地址信息，取得省份
                            JSONObject expendJson = JSONObject.parseObject(row.getCell(colIndex).getStringCellValue());
                            JSONObject addressJson = expendJson.getJSONObject("address");
                            provinceId = Integer.valueOf((String) addressJson.get("provinceId"));
                            provinceName = (String) addressJson.get("provinceName");
                            provinceMap.put(provinceId, provinceName);
                            provinceNum = provinceIdList.get(provinceId);
                            provinceIdList.set(provinceId, ++provinceNum);
                            realOrder = false;
                        }
                        if (colIndex == 10) {// 记录sku数量，用于筛选热门sku
                            sku = (long) row.getCell(colIndex).getNumericCellValue();
                            if (skuMap.get(sku) != null) {
                                skuNum = skuMap.get(sku);
                                skuMap.put(sku, ++skuNum);
                            } else {
                                skuMap.put(sku, 1);
                            }
                        }
                        if (colIndex == 11) {// 记录sku商品文案
                            name = row.getCell(colIndex).getStringCellValue();
                            skuNameMap.put(sku, name);

                            if (maxPayPrice && !topPriceOrderIds.contains(ordId)) {// 筛选最大价格订单
                                buildJdSearchOrder(jdSearchOrder, ordId, orderDate, freight, price, coupon, provinceId, provinceName, sku, name);

                                if (topPriceOrders.size() < 10) {//直接放
                                    topPriceOrders.add(jdSearchOrder);
                                    if (topPriceOrders.size() > 1) {
                                        Collections.sort(topPriceOrders, new PriceComparator());
                                    }
                                } else {// 去掉最小的，放入新元素并排序
                                    topPriceOrders.remove(0);
                                    topPriceOrders.add(0, jdSearchOrder);
                                    Collections.sort(topPriceOrders, new PriceComparator());
                                    maxPriceOrder = topPriceOrders.get(0);
                                }
                                topPriceOrderIds.add(ordId);
                                maxPayPrice = false;
                            }
                            if (maxCoupon && !topCouponOrderIds.contains(ordId)) {// 筛选最大电子券
                                buildJdSearchOrder(jdSearchOrder, ordId, orderDate, freight, price, coupon, provinceId, provinceName, sku, name);

                                if (topCouponOrders.size() < 10) {//直接放
                                    topCouponOrders.add(jdSearchOrder);
                                    Collections.sort(topCouponOrders, new CouponComparator());
                                } else {// 去掉最小的，放入新元素并排序
                                    topCouponOrders.remove(0);
                                    topCouponOrders.add(0, jdSearchOrder);
                                    Collections.sort(topCouponOrders, new CouponComparator());
                                    maxCouponOrder = topCouponOrders.get(0);
                                }
                                topCouponOrderIds.add(ordId);
                                maxCoupon = false;
                            }
                        }
                    }
                }
            }
            List<List<List<Object>>> data = new ArrayList<>();

            // 找出前五名省份
            List<List<Object>> topProvinces = new ArrayList<>();
            sortAndGetProvince(topProvinces, provinceIdList, provinceMap, 5);

            // 排序找出sku前十名
            List<List<Object>> topSkus = new ArrayList<>();
            sortAndGetSku(topSkus, skuMap, skuNameMap, 10);

            //处理订单金额最大前十名
            List<List<Object>> topPrices = new ArrayList<>();
            treatTopPriceOrders(topPriceOrders, topPrices);

            //处理订单金额最大前十名
            List<List<Object>> topCoupons = new ArrayList<>();
            treatTopPriceOrders(topCouponOrders, topCoupons);

            data.add(0, topProvinces);
            data.add(1, topSkus);
            data.add(2, topPrices);
            data.add(3, topCoupons);
            return data;
        }

        private void buildJdSearchOrder(JDSearchOrderVo jdSearchOrder, Long ordId, Integer orderDate, Integer freight,
                                        Integer price, Integer coupon, Integer provinceId, String provinceName, Long sku, String name) {
            jdSearchOrder.setJdOrderId(ordId);
            jdSearchOrder.setOrderDate(orderDate);
            jdSearchOrder.setFreight(freight);
            jdSearchOrder.setPayPrice(price);
            jdSearchOrder.setCompanyPaymony(coupon);
            jdSearchOrder.setProvinceId(provinceId);
            jdSearchOrder.setProvinceName(provinceName);
            jdSearchOrder.setSku(sku);
            jdSearchOrder.setName(name);
        }

        private void treatTopPriceOrders(List<JDSearchOrderVo> topPriceOrders, List<List<Object>> topPrices) {
            List<Object> singleOrder = null;
            int i = 0;
            for (i = 9; i >= 0; i--) {
                singleOrder = new ArrayList<>();
                singleOrder.add(0, topPriceOrders.get(i).getJdOrderId());
                singleOrder.add(1, topPriceOrders.get(i).getOrderDate());
                singleOrder.add(2, topPriceOrders.get(i).getFreight());
                singleOrder.add(3, topPriceOrders.get(i).getPayPrice());
                singleOrder.add(4, topPriceOrders.get(i).getCompanyPaymony());
                singleOrder.add(5, topPriceOrders.get(i).getProvinceName());
                singleOrder.add(6, topPriceOrders.get(i).getSku());
                singleOrder.add(7, topPriceOrders.get(i).getName());
                topPrices.add(9 - i, singleOrder);
            }
        }

        private void sortAndGetSku(List<List<Object>> topSkus, Map<Long, Integer> skuMap,
                                   Map<Long, String> skuNameMap, int topNum) {
            List<Object> singleSku = null;
            while (topNum > 0) {
                singleSku = new ArrayList<>();
                long[] temp = new long[2];// {sku, 单数}
                for (Map.Entry<Long, Integer> entry : skuMap.entrySet()) {
                    if (entry.getValue() == null) {
                        continue;
                    }
                    if (entry.getValue() > temp[1]) {
                        temp[1] = entry.getValue();
                        temp[0] = entry.getKey();
                    }
                }
                singleSku.add(0, temp[0]);
                singleSku.add(1, temp[1]);
                singleSku.add(2, skuNameMap.get(temp[0]));
                topSkus.add(singleSku);
                skuMap.remove(temp[0]);
                topNum--;
            }
        }

        private void sortAndGetProvince(List<List<Object>> topProvinces, List<Integer> provinceIdList, Map<Integer, String> provinceMap, int topNum) {
            List<Object> province = null;
            while (topNum > 0) {
                province = new ArrayList<>();
                int[] temp = new int[2];// {值, 索引}
                for (int j = 0; j < provinceIdList.size(); j++) {
                    if (provinceIdList.get(j) == null) {
                        continue;
                    }
                    if (provinceIdList.get(j) > temp[0]) {
                        temp[0] = provinceIdList.get(j);
                        temp[1] = j;
                    }
                }
                province.add(0, 6-topNum);
                province.add(1, provinceMap.get(temp[1]));
                province.add(2, temp[0]);
                topProvinces.add(5-topNum, province);
                provinceIdList.remove(temp[1]);
                topNum--;
            }
        }

    }

    static class PriceComparator implements Comparator<Object> {
        public int compare(Object object1, Object object2) {
            JDSearchOrderVo o1 = (JDSearchOrderVo) object1;
            JDSearchOrderVo o2 = (JDSearchOrderVo) object2;
            return new Integer(o1.getPayPrice()).compareTo(new Integer(o2.getPayPrice()));
        }
    }

    static class CouponComparator implements Comparator<Object> {
        public int compare(Object object1, Object object2) {
            JDSearchOrderVo o1 = (JDSearchOrderVo) object1;
            JDSearchOrderVo o2 = (JDSearchOrderVo) object2;
            return new Integer(o1.getCompanyPaymony()).compareTo(new Integer(o2.getCompanyPaymony()));
        }
    }

    public static void main(String[] args) {
//		Thread threadWeekly = new Thread(new WeeklySale());
//		threadWeekly.start();
        treateExcelData();

    }

}
