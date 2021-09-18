import bean.Class;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.*;
import java.util.*;


public class SummaryClass {


    public static void main(String[] args) throws IOException {
        SummaryClass summary = new SummaryClass();
        String dirPath = "./file/";
        File files = new File(dirPath);
        File[] fileList = files.listFiles();
        if (fileList != null) {
            for (File file : fileList) {
                if (file.getName().contains(".docx")) {
                    System.out.println("Start handle word file:" + file.getName());
                    ArrayList<Class> classes = summary.HandelWord(dirPath + file.getName());
                    String[] titles = summary.GetTitles();
                    String excelFileName = file.getName().replace(".docx", ".xlsx");
                    summary.ExportExcel(classes, titles, excelFileName);
                    System.out.println("generate excel file:" + excelFileName);

                }
            }
        }
    }

    public ArrayList<Class> HandelWord(String filePath) {
        try {
            ArrayList<Class> classes = new ArrayList<>();
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
                        Class cls = new Class();
                        if (cells.size() < 2)
                            continue;
                        for (int j = 0; j < cells.size(); j++) {
                            XWPFTableCell cell = cells.get(j);
                            if (j == 0) {
                                cls.name = cell.getText();
                            } else {
                                String content = cell.getText();
                                if (!Objects.equals(content, "")) {
                                    // 根据"："切割
                                    String[] starList = cell.getText().split("：");
                                    for (int starIndex = 1; starIndex < starList.length; starIndex++) {
                                        // 获取当前事件分类
                                        String star = starList[starIndex - 1].substring(starList[starIndex - 1].length() - 3);
                                        //当前事件总分
                                        float score = 0;
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

                                        if (star.equals("道德星")) {
                                            cls.moral = score;
                                        } else if (star.equals("阅读星")) {
                                            cls.read = score;
                                        } else if (star.equals("智慧星")) {
                                            cls.wisdom = score;
                                        } else if (star.equals("健康星")) {
                                            cls.health = score;
                                        } else if (star.equals("艺术星")) {
                                            cls.art = score;
                                        } else if (star.equals("实践星")) {
                                            cls.practice = score;
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
                        classes.add(cls);
                    }
                }
            }
            return classes;
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

    public float HandleEvent(String str) {
        String[] events = str.split(" ");
        List<String> list = new ArrayList<>();
        for (String event : events) {
            if (!event.equals("")) {
                list.add(event);
            }
        }
        float score = 0;
        for (String event : list) {
            score += EventScore(event);
        }
        return score;
    }

    public float EventScore(String event) {
        StringBuilder numStr = new StringBuilder();
        float num = 0;
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
                            num += Float.parseFloat(numStr.toString());
                            numStr = new StringBuilder();
                        }
                    }
                }
            } else {
                if (numStr.length() != 0) {
                    if (b == '.' || (b >= '0' && b <= '9')) {
                        numStr.append(b);

                    }
                    num += Float.parseFloat(numStr.toString());
                    numStr = new StringBuilder();
                }
            }
        }
        return num;
    }

    public void ExportExcel(ArrayList<Class> classes, String[] titles, String filename) throws IOException {
        String xlsxPath = "./file/" + filename;
        Workbook workBook = new XSSFWorkbook();
        OutputStream fos = null;
        try {
            Sheet sheet = workBook.createSheet("sheet1");

            sheet.setDefaultColumnWidth(10);
            Row row = sheet.createRow((int) 0);
            CellStyle style = workBook.createCellStyle();
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

            Cell cell = null;
            for (int i = 0; i < titles.length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(titles[i]);
                cell.setCellStyle(style);
            }

            for (int i = 0; i < classes.size(); i++) {
                row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(classes.get(i).name);
                row.createCell(1).setCellValue(classes.get(i).moral);
                row.createCell(2).setCellValue(classes.get(i).read);
                row.createCell(3).setCellValue(classes.get(i).wisdom);
                row.createCell(4).setCellValue(classes.get(i).health);
                row.createCell(5).setCellValue(classes.get(i).art);
                row.createCell(6).setCellValue(classes.get(i).practice);
                switch (classes.get(i).star) {
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
            fos = new FileOutputStream(xlsxPath);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        workBook.write(fos);
        fos.close();
    }
}



