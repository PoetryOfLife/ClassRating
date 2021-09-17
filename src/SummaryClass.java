import bean.Class;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;


public class SummaryClass {


    public static void main(String[] args) {
        SummaryClass summary = new SummaryClass();
        String dirPath = "./file/";
        File files = new File(dirPath);
        File[] fileList = files.listFiles();
        if (fileList != null) {
            for (File file : fileList) {
                if (file.getName().contains(".docx")) {
                    ArrayList<Class> classes = summary.HandelWord(dirPath + file.getName());

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
                        for (int j = 0; j < cells.size(); j++) {
                            XWPFTableCell cell = cells.get(j);
                            if (j == 0) {
                                cls.name = cell.getText();
                                System.out.println("class:" + cell.getText());
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

                                        if (score >= 0) {
                                            cls.star++;
                                        }
                                    }
                                }
                            }
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
        String[] scores;
        if (event.contains("+")) {
            scores = event.split("\\+");
            return Float.parseFloat(scores[1]);
        } else if (event.contains("-")) {
            scores = event.split("\\-");
            return -Float.parseFloat(scores[1]);
        }
        return 0;
    }

    public void ExportExcel(ArrayList<Class> classes) {

    }
}



