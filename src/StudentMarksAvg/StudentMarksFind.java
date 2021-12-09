/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package StudentMarksAvg;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashSet;
import java.util.Iterator;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class StudentMarksFind {

    public static double find(int id, int marks) {
        double relevantMark = 0;
        try {

            boolean state = false;
            File file = new File("C:\\Users\\sajeevan\\Desktop\\Student Marks Quer Form\\StudentMarksAverage\\src\\DataSet\\Assign2_student_results.xlsx");
            FileInputStream fis = new FileInputStream(file);

            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);

            if (marks == 6) {
                double assignmentMark = 0;
                double quizeMark = 0;
                double midTermMark = 0;
                double projectMark = 0;
                double finalExamMark = 0;
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            if (id == row.getCell(0).getNumericCellValue()) {
                                assignmentMark = row.getCell(1).getNumericCellValue();
                                quizeMark = row.getCell(2).getNumericCellValue();
                                midTermMark = row.getCell(3).getNumericCellValue();
                                projectMark = row.getCell(4).getNumericCellValue();
                                finalExamMark = row.getCell(5).getNumericCellValue();
                            }

                        }
                    }
                }
                relevantMark = assignmentMark + quizeMark + midTermMark + projectMark + finalExamMark;
            } else {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            if (id == row.getCell(0).getNumericCellValue()) {
                                relevantMark = (double) row.getCell(marks).getNumericCellValue();

                            }

                        }
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

//        System.out.println(relevantMark);
        return relevantMark;

    }
}
