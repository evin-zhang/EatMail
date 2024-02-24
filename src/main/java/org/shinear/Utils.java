package org.shinear;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.mail.BodyPart;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.internet.MimeMessage;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Utils {


    // 定义一个自定义的方法，将邮件内容按照格式转换成 Excel 文档，并保存到目标文件中
    public static void convertEmailToExcel(String emailFilePath, String excelFilePath)  {
        try {
            // 读取邮件
            String mailContent = parseMail(emailFilePath);


            // 使用正则表达式匹配数据
            Pattern pattern = Pattern.compile("(.*)：(.*)");
            //Matcher matcher = pattern.matcher(mailContent);
            String[] splits = mailContent.split("充电桩客服大橙子🍊 \\d{2}:\\d{2}");

            // 存储数据
            List<Map<String, String>> records = new ArrayList<>();
            //Map<String,String> record = new LinkedHashMap<>();
            List<String> columnOrder = Arrays.asList("安装完成时间", "客户姓名", "车辆品牌", "增项价格", "增项内容", "物料总用量", "安装人员姓名", "出发地", "安装地");
            /**
            while (matcher.find()) {
                String key = matcher.group(1).trim();
                String value = matcher.group(2).trim();
                if ("安装完成时间".equals(key) && !record.isEmpty()) {
                    records.add(record);
                    record = new LinkedHashMap<>();
                }
                record.put(key, value);
            }
            records.add(record);  // 添加最后一条记录
            */
            // 按照每条记录进行分割
            for (String split : splits) {
                Matcher matcher = pattern.matcher(split);
                Map<String, String> record = new LinkedHashMap<>();
                while (matcher.find()) {
                    String key = matcher.group(1).trim();
                    String value = matcher.group(2).trim();
                    record.put(key, value);
                }
                if (!record.isEmpty()) {
                    records.add(record);
                }
            }

            // 创建Excel文件
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Data");
            // 创建字体样式
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);

            // 添加表头
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < columnOrder.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columnOrder.get(i));
                cell.setCellStyle(headerCellStyle);  // 设置字体样式
            }
            // 填充数据
            int rowNum = 1;
            for (Map<String, String> rec : records) {
                Row row = sheet.createRow(rowNum++);
                for (String column : columnOrder) {
                    String value = rec.get(column);
                    if (value != null) {
                        Cell cell = row.createCell(columnOrder.indexOf(column));
                        cell.setCellValue(value);
                    }
                }
            }



            // 保存 Excel 文档
            String timestamp = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
            String outputFilePath = excelFilePath + File.separator + new File(emailFilePath).getName() + "_" + timestamp + ".xlsx";
            FileOutputStream out = new FileOutputStream(outputFilePath);
            workbook.write(out);
            out.close();
            workbook.close();

            System.out.println("Excel 文件已保存到: " + outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static String parseMail(String emailFilePath) throws MessagingException, IOException {
        Session session = Session.getDefaultInstance(new Properties());
        InputStream inMsg = new FileInputStream(emailFilePath);
        MimeMessage message = new MimeMessage(session, inMsg);

        // 获取邮件正文
        Object content = message.getContent();
        String bodyText = "";
        if (content instanceof String) {
            bodyText = (String) content;
        } else if (content instanceof Multipart) {
            Multipart multipart = (Multipart) content;
            for (int i = 0; i < multipart.getCount(); i++) {
                BodyPart bodyPart = multipart.getBodyPart(i);
                if (bodyPart.isMimeType("text/plain")) {
                    bodyText = (String) bodyPart.getContent();
                    break;
                }
            }
        }
        return bodyText;
    }

}
