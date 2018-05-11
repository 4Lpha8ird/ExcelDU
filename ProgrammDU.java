/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package programmdu; // Папка с программой

import java.io.FileInputStream; // Чтение
import java.io.FileNotFoundException; // Исключение
import java.io.FileOutputStream; // Запись
import java.io.IOException; // Исключение
import org.apache.poi.hssf.usermodel.HSSFWorkbook; // Распознование Excel файла формата .xls
import org.apache.poi.ss.usermodel.Cell; // Ячейки
import org.apache.poi.ss.usermodel.DateUtil; // Ячейки формата "дата"
import org.apache.poi.ss.usermodel.*; // Работа над документом

/**
 *
 * @author Sergey
 */
public class ProgrammDU { // Создание класса

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, FileNotFoundException { // Главный метод

        FileInputStream fis = new FileInputStream("C:/Users/$ergey$/Desktop/Программа/Dano.xls"); // Путь к файлу и его чтение

        Workbook wb = new HSSFWorkbook(fis); // Работа с 1 файлом ("Dano") Объявляем поток
        Workbook wb2 = new HSSFWorkbook(); // Работа с 2 файлом ("Otvet") 

        Sheet newSheet = wb2.createSheet("123"); // Создаем лист во 2ом файле 

        String a1 = wb.getSheetAt(0).getRow(0).getCell(1).getStringCellValue(); // Чтение данных из Excel в программу (В формате String)
        String b1 = wb.getSheetAt(0).getRow(1).getCell(1).getStringCellValue();
        String c1 = wb.getSheetAt(0).getRow(2).getCell(1).getStringCellValue();
        String d1 = wb.getSheetAt(0).getRow(3).getCell(1).getStringCellValue();
        String f1 = wb.getSheetAt(0).getRow(4).getCell(1).getStringCellValue();
        String g1 = wb.getSheetAt(0).getRow(5).getCell(1).getStringCellValue();
        String x01 = wb.getSheetAt(0).getRow(6).getCell(1).getStringCellValue();
        String y01 = wb.getSheetAt(0).getRow(7).getCell(1).getStringCellValue();
        String h1001 = wb.getSheetAt(0).getRow(8).getCell(1).getStringCellValue();

        double a = Double.parseDouble(a1);// Переводим переменные формата String в Double
        double b = Double.parseDouble(b1);
        double c = Double.parseDouble(c1);
        double d = Double.parseDouble(d1);
        double f = Double.parseDouble(f1);
        double g = Double.parseDouble(g1);
        double x0 = Double.parseDouble(x01);
        double y0 = Double.parseDouble(y01);
        double h = Double.parseDouble(h1001);
        double[] x = new double[50];
        double[] y = new double[50];
        x[0] = x0;
        y[0] = y0;
        double dy = d * Math.pow(Math.E, f * (y[0])) + Math.log(g * (x[0]));
        double hdy = (h / 10) * dy;
        double k1; // Коэффициенты Рунге-Кутта
        double k2;
        double k3;
        double k4;
        k1 = d * Math.pow(Math.E, f * (y[0])) + Math.log(g * (x[0]));
        k2 = d * Math.pow(Math.E, f * (y[0] + (h * k1) / 20)) + Math.log(g * (x[0] + h / 20));
        k3 = d * Math.pow(Math.E, f * (y[0] + (h * k2) / 20)) + Math.log(g * (x[0] + h / 20));
        k4 = d * Math.pow(Math.E, f * (y[0] + (h * k3) / 10)) + Math.log(g * (x[0] + h / 10));
        double dy0 = (h / 60) * (k1 + 2 * k2 + 2 * k3 + k4);
        Row row12 = newSheet.createRow(9); // Создаем строку
        row12.createCell(5).setCellValue("x"); // Создаем столбец. Получаем ячейку и вписываем в нее значение.
        row12.createCell(6).setCellValue(x[0]);
        row12.createCell(7).setCellValue("y");
        row12.createCell(8).setCellValue(y[0]);
        row12.createCell(3).setCellValue("Эйлер");
        row12.createCell(4).setCellValue("Ответ");
        Row row13 = newSheet.createRow(23);
        row13.createCell(5).setCellValue("x");
        row13.createCell(6).setCellValue(x[0]);
        row13.createCell(7).setCellValue("y");
        row13.createCell(8).setCellValue(y[0]);
        row13.createCell(3).setCellValue("Рунге-Кутт 4-го порядка");
        row13.createCell(4).setCellValue("Ответ");
        Row row14 = newSheet.createRow(37);
        row14.createCell(3).setCellValue("Рунге-Кутт 3-го порядка");
        row14.createCell(4).setCellValue("Ответ");
        row14.createCell(5).setCellValue("x");
        row14.createCell(6).setCellValue(x[0]);
        row14.createCell(7).setCellValue("y");
        row14.createCell(8).setCellValue(y[0]);

        for (int i = 1; i <= 10; i++) { // метод Эйлера
            x[i] = x[i - 1] + h / 10;
            y[i] = y[i - 1] + hdy;
            dy = d * Math.pow(Math.E, f * (y[i])) + Math.log(g * (x[i]));
            hdy = h / 10 * dy;
            Row row11 = newSheet.createRow(i + 9);
            row11.createCell(5).setCellValue("x");
            row11.createCell(6).setCellValue(x[i]);
            row11.createCell(7).setCellValue("y");
            row11.createCell(8).setCellValue(y[i]);
            FileOutputStream fileOut = new FileOutputStream("C:/Users/$ergey$/Desktop/Программа/Otvet.xls");
            wb2.write(fileOut);
            fis.close();
        }

        h = Double.parseDouble(h1001);
        dy = d * Math.pow(Math.E, f * (y[0])) + Math.log(g * (x[0]));
        hdy = h * dy;

        for (int i = 1; i <= 10; i++) { // метод Рунге-Кутта 4-го порядка
            x[i] = x[i - 1] + h / 10;
            y[i] = y[i - 1] + dy0;
            k1 = d * Math.pow(Math.E, f * (y[i])) + Math.log(g * (x[i]));
            k2 = d * Math.pow(Math.E, f * (y[i] + (h * k1) / 20)) + Math.log(g * (x[i] + h / 20));
            k3 = d * Math.pow(Math.E, f * (y[i] + (h * k2) / 20)) + Math.log(g * (x[i] + h / 20));
            k4 = d * Math.pow(Math.E, f * (y[i] + h * k3 / 10)) + Math.log(g * (x[i] + h / 10));
            dy0 = (h / 60) * (k1 + 2 * k2 + 2 * k3 + k4);
            Row row11 = newSheet.createRow(i + 23);
            row11.createCell(5).setCellValue("x");
            row11.createCell(6).setCellValue(x[i]);
            row11.createCell(7).setCellValue("y");
            row11.createCell(8).setCellValue(y[i]);
            FileOutputStream fileOut = new FileOutputStream("C:/Users/$ergey$/Desktop/Программа/Otvet.xls");
            wb2.write(fileOut);
            fis.close();
        }

        h = Double.parseDouble(h1001);
        dy = d * Math.pow(Math.E, f * (y[0])) + Math.log(g * (x[0]));
        hdy = h * dy;
        k1 = d * Math.pow(Math.E, f * (y[0])) + Math.log(g * (x[0]));
        k2 = d * Math.pow(Math.E, f * (y[0] + (h * k1) / 20)) + Math.log(g * (x[0] + h / 20));
        k3 = d * Math.pow(Math.E, f * (y[0] + (h * k2) / 20)) + Math.log(g * (x[0] + h / 20));
        k4 = d * Math.pow(Math.E, f * (y[0] + (h * k3) / 10)) + Math.log(g * (x[0] + h / 10));

        for (int i = 1; i <= 10; i++) { // метод Рунге-Кутта 3-го порядка
            y[i] = y[i - 1] + dy0;
            x[i] = x[i - 1] + h / 10;
            k1 = h * (d * Math.pow(Math.E, f * (y[i])) + Math.log(g * (x[i])));
            k2 = h * (d * Math.pow(Math.E, f * (y[i] + k1 / 2)) + Math.log(g * (x[i] + h / 20)));
            k3 = h * (d * Math.pow(Math.E, f * (y[i] + 2 * k2 - k1)) + Math.log(g * (x[i] + h / 10)));
            dy0 = (k1 + 4 * k2 + k3) / 60;
            Row row11 = newSheet.createRow(i + 37);
            row11.createCell(5).setCellValue("x");
            row11.createCell(6).setCellValue(x[i]);
            row11.createCell(7).setCellValue("y");
            row11.createCell(8).setCellValue(y[i]);
            FileOutputStream fileOut = new FileOutputStream("C:/Users/$ergey$/Desktop/Программа/Otvet.xls");
            wb2.write(fileOut);
            fis.close();
        }

        h = Double.parseDouble(h1001) / 10;
        Row row0 = newSheet.createRow(0);
        row0.createCell(0).setCellValue("a");
        row0.createCell(1).setCellValue(a);
        Row row1 = newSheet.createRow(1);
        row1.createCell(0).setCellValue("b");
        row1.createCell(1).setCellValue(b);
        Row row2 = newSheet.createRow(2);
        row2.createCell(0).setCellValue("c");
        row2.createCell(1).setCellValue(c);
        Row row3 = newSheet.createRow(3);
        row3.createCell(0).setCellValue("d");
        row3.createCell(1).setCellValue(d);
        Row row4 = newSheet.createRow(4);
        row4.createCell(0).setCellValue("f");
        row4.createCell(1).setCellValue(f);
        Row row5 = newSheet.createRow(5);
        row5.createCell(0).setCellValue("g");
        row5.createCell(1).setCellValue(g);
        Row row6 = newSheet.createRow(6);
        row6.createCell(0).setCellValue("x0");
        row6.createCell(1).setCellValue(x0);
        Row row7 = newSheet.createRow(7);
        row7.createCell(0).setCellValue("y0");
        row7.createCell(1).setCellValue(y0);
        Row row8 = newSheet.createRow(8);
        row8.createCell(0).setCellValue("h");
        row8.createCell(1).setCellValue(h);
        FileOutputStream fileOut = new FileOutputStream("C:/Users/$ergey$/Desktop/Программа/Otvet.xls"); // Создаем файл 
        wb2.write(fileOut); // Записываем в него вышесозданные ячейки
        fis.close(); // Закрываем поток чиения

    }

    public static String gettext(Cell cell) { // Тонкости библиотеки poi, связанные с чтением ячеек разных форматов. Данный метод читает, а затем переводит переменные любого типа в String
        String result = "";
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                System.out.println(cell.getRichStringCellValue().getString());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = cell.getDateCellValue().toString();
                } else {
                    result = Double.toString(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                result = Boolean.toString(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                result = cell.getCellFormula().toString();
                break;
            default:
                break;
        }
        return result;
    }

}
