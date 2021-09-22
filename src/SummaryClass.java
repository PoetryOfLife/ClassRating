import bean.ClassRating;
import bean.RoutineInspection;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import services.FormatData;

import java.io.*;
import java.text.DecimalFormat;
import java.util.*;


public class SummaryClass {


    public static void main(String[] args) throws IOException, InterruptedException {
        SummaryClass summary = new SummaryClass();
        FormatData ft = new FormatData();
        String dirPath = "./file/";
        File files = new File(dirPath);
        File[] fileList = files.listFiles();

        ArrayList<RoutineInspection> routineInspections = new ArrayList<>();
        if (fileList != null) {

            for (File file : fileList) {
                if (file.getName().contains(".docx")) {
                    System.out.println("Start handle word file:" + file.getName());
                    routineInspections.add(summary.HandelWord(dirPath, file.getName()));
                }
            }

            routineInspections.add(ft.SummaryInspection(routineInspections));

            String[] titles = summary.GetTitles();
            summary.ExportExcel(routineInspections, titles, "少先队日常规检查详情反馈.xlsx");

        } else {
            System.out.println("no file!");
        }
        Thread.sleep(1000);


    }

    public RoutineInspection HandelWord(String dirPath, String fileName) {
        String filePath = dirPath + fileName;
        try {
            ArrayList<ClassRating> classRatings = new ArrayList<>();
            FileInputStream in = new FileInputStream(filePath);
            if (filePath.toLowerCase().endsWith("docx")) {
                XWPFDocument xwpf = new XWPFDocument(in);
                Iterator<XWPFTable> it = xwpf.getTablesIterator();
                while (it.hasNext()) {
                    XWPFTable table = it.next();
                    List<XWPFTableRow> rows = table.getRows();
                    for (int i = 2; i < rows.size(); i++) {
                        XWPFTableRow row = rows.get(i);
                        List<XWPFTableCell> cells = row.getTableCells();
                        ClassRating cls = new ClassRating();
                        if (cells.size() < 2)
                            continue;
                        for (int j = 0; j < cells.size(); j++) {
                            XWPFTableCell cell = cells.get(j);
                            if (j == 0) {
                                cls.className = cell.getText();
                            } else {
                                String content = cell.getText();
                                if (!Objects.equals(content, "")) {
                                    // 根据"："切割
                                    String[] starList = cell.getText().split("：");
                                    for (int starIndex = 1; starIndex < starList.length; starIndex++) {
                                        // 获取当前事件分类
                                        String star = starList[starIndex - 1].substring(starList[starIndex - 1].length() - 3);
                                        //当前事件总分
                                        double score;
                                        // 记录分类分数
                                        if (starList.length > 2) {
                                            if (starIndex != starList.length - 1) {
                                                String events = starList[starIndex].substring(0, starList[starIndex].length() - 3);
                                                score = HandleEvent(events);
                                            } else {
                                                score = HandleEvent(starList[starIndex]);
                                            }
                                        } else {
                                            score = HandleEvent(starList[starIndex]);
                                        }

                                        switch (star) {
                                            case "道德星":
                                                cls.moral = score;
                                                break;
                                            case "阅读星":
                                                cls.read = score;
                                                break;
                                            case "智慧星":
                                                cls.wisdom = score;
                                                break;
                                            case "健康星":
                                                cls.health = score;
                                                break;
                                            case "艺术星":
                                                cls.art = score;
                                                break;
                                            case "实践星":
                                                cls.practice = score;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        if (cls.moral >= 0) {
                            cls.star++;
                        }
                        if (cls.read >= 0) {
                            cls.star++;
                        }
                        if (cls.wisdom >= 0) {
                            cls.star++;
                        }
                        if (cls.health >= 0) {
                            cls.star++;
                        }
                        if (cls.art >= 0) {
                            cls.star++;
                        }
                        if (cls.practice >= 0) {
                            cls.star++;
                        }
                        classRatings.add(cls);
                    }
                }
            }
            return new RoutineInspection(classRatings, fileName);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    public String[] GetTitles() {
        String[] titles = new String[10];
        titles[0] = "班级";
        titles[1] = "道德星";
        titles[2] = "阅读星";
        titles[3] = "智慧星";
        titles[4] = "健康星";
        titles[5] = "艺术星";
        titles[6] = "实践星";
        titles[7] = "星级班级";
        return titles;
    }

    public double HandleEvent(String str) {
        String[] events = str.split(" ");
        List<String> list = new ArrayList<>();
        for (String event : events) {
            if (!event.equals("")) {
                list.add(event);
            }
        }
        double score = 0;
        for (String event : list) {
            score += EventScore(event);
        }
        return score;
    }

    public double EventScore(String event) {
        StringBuilder numStr = new StringBuilder();
        double num = 0;
        for (int i = 0; i < event.length(); i++) {
            char b = event.charAt(i);
            if (i != event.length() - 1) {
                if (b == '+' || b == '-') {
                    numStr.append(b);
                } else {
                    if (numStr.length() != 0) {
                        if (b == '.' || (b >= '0' && b <= '9')) {
                            numStr.append(b);
                        } else {
                            num += Double.parseDouble(numStr.toString());
                            numStr = new StringBuilder();
                        }
                    }
                }
            } else {
                if (numStr.length() != 0) {
                    if (b == '.' || (b >= '0' && b <= '9')) {
                        numStr.append(b);

                    }
                    num += Double.parseDouble(numStr.toString());
                    numStr = new StringBuilder();
                }
            }
        }
        return num;
    }

    public void ExportExcel(ArrayList<RoutineInspection> routineInspections, String[] titles, String filename) throws IOException {
        String xlsxPath = "./file/" + filename;
        Workbook workBook = new XSSFWorkbook();
        DecimalFormat df = new DecimalFormat("0.00");
        OutputStream fos = null;

        try {

            for (RoutineInspection routineInspection : routineInspections) {
                Sheet sheet = workBook.createSheet(routineInspection.fileName);
                sheet.setDefaultColumnWidth(10);
                Row row = sheet.createRow(0);
                CellStyle style = workBook.createCellStyle();
                style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

                Cell cell;
                for (int i = 0; i < titles.length; i++) {
                    cell = row.createCell(i);
                    cell.setCellValue(titles[i]);
                    cell.setCellStyle(style);
                }

                ArrayList<ClassRating> classRatings = routineInspection.cr;

                for (int i = 0; i < classRatings.size(); i++) {
                    row = sheet.createRow(i + 1);
                    row.createCell(0).setCellValue(classRatings.get(i).className);
//                    row.createCell(1).setCellValue(df.format(classRatings.get(i).moral));
//                    row.createCell(2).setCellValue(df.format(classRatings.get(i).read));
//                    row.createCell(3).setCellValue(df.format(classRatings.get(i).wisdom));
//                    row.createCell(4).setCellValue(df.format(classRatings.get(i).health));
//                    row.createCell(5).setCellValue(df.format(classRatings.get(i).art));
//                    row.createCell(6).setCellValue(df.format(classRatings.get(i).practice));

                    row.createCell(1).setCellValue(classRatings.get(i).moral);
                    row.createCell(2).setCellValue(classRatings.get(i).read);
                    row.createCell(3).setCellValue(classRatings.get(i).wisdom);
                    row.createCell(4).setCellValue(classRatings.get(i).health);
                    row.createCell(5).setCellValue(classRatings.get(i).art);
                    row.createCell(6).setCellValue(classRatings.get(i).practice);

                    switch (classRatings.get(i).star) {
                        case 0:
                            row.createCell(7).setCellValue("零星班级");
                            break;
                        case 1:
                            row.createCell(7).setCellValue("一星班级");
                            break;
                        case 2:
                            row.createCell(7).setCellValue("二星班级");
                            break;
                        case 3:
                            row.createCell(7).setCellValue("三星班级");
                            break;
                        case 4:
                            row.createCell(7).setCellValue("四星班级");
                            break;
                        case 5:
                            row.createCell(7).setCellValue("五星班级");
                            break;
                        case 6:
                            row.createCell(7).setCellValue("六星班级");
                            break;
                    }
                }
            }

            fos = new FileOutputStream(xlsxPath);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        workBook.write(fos);
        assert fos != null;
        fos.close();
    }
}



