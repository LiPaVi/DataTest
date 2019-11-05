import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.Scanner;

public class HumanData {
    public static void main(String[] args) throws Exception{


        Scanner in = new Scanner(System.in);
        System.out.print("Введите число от 1 до 30: ");
        int rownum = in.nextInt();
        in.close();

//        boolean sex;
//
//        String names[];
//        String surnames[];
//        String patronymic[];
//
//        Date birth;
//        int age;
//
//        int index;
//        String country = "Россия";
//        String streets[];
//        int houseNumber;
//        int flatNumber;

//        Struct Human;


        ArrayList<String> femaleNames = getArrayFromFile("src/main/resources/FemaleNames.txt");
        ArrayList<String> maleNames = getArrayFromFile("src/main/resources/MaleNames.txt");
        ArrayList<String> surnames = getArrayFromFile("src/main/resources/Surnames.txt");
        ArrayList<String> femalePatronymic = getArrayFromFile("src/main/resources/FemalePatronymic.txt");
        ArrayList<String> malePatronymic = getArrayFromFile("src/main/resources/MalePatronymic.txt");
        ArrayList<String> streets = getArrayFromFile("src/main/resources/Streets.txt");
        ArrayList<String> cities = getArrayFromFile("src/main/resources/Cities.txt");



        HSSFWorkbook workBook = new HSSFWorkbook();
        HSSFSheet sheet = workBook.createSheet("Тестовые данные");

        // создаем шрифт
        HSSFFont font = workBook.createFont();
        // указываем, что хотим его видеть жирным
        font.setBold(true);
        // создаем стиль для ячейки
        HSSFCellStyle style = workBook.createCellStyle();
        // и применяем к этому стилю жирный шрифт
        style.setFont(font);

        int i =0;

        while (i <= rownum) {
            sheet.createRow(i);
            i++;
        }


        fillTheColumn(sheet, rownum, 0, maleNames, "Имя");
        fillTheColumn(sheet, rownum, 1, surnames, "Фамилия");
        fillTheColumn(sheet, rownum, 2, malePatronymic, "Отчество");
        fillTheColumn(sheet, rownum, 3, cities, "Город");
        fillTheColumn(sheet, rownum, 4, streets, "Улица");

        // Создаем файл
        File file = new File("src/main/resources/data.xls");
//        file.getParentFile().mkdirs();

        FileOutputStream outFile = new FileOutputStream(file);
        workBook.write(outFile);
        System.out.println("Файл создан. Путь: " + file.getAbsolutePath());
        workBook.close();

    }

    private static int randomNumber(int max){
        return (int) (Math.random() * max);
    }

    private static void fillTheColumn(HSSFSheet sheet, int rownum, int columnnum, ArrayList<String> list, String columnName){
        String cellValue;
        Cell cell;
        Row row;

        row = sheet.getRow(0);
        cell = row.createCell(columnnum, CellType.STRING);
        cell.setCellValue(columnName);

        int i = 1;
        while (i <= rownum) {
            cellValue = list.get(randomNumber(list.size()));
            row = sheet.getRow(i);
            cell = row.createCell(columnnum, CellType.STRING);
            cell.setCellValue(cellValue);
            i += 1;
        }
    }

    private static ArrayList<String> getArrayFromFile(String path) throws Exception{
        ArrayList<String> array = new ArrayList<String>();

        FileReader cityFile = new FileReader(path);
        Scanner scan = new Scanner(cityFile);

        while (scan.hasNextLine()){
            array.add(scan.nextLine());
        }

        cityFile.close();
        return array;
    }

    private static ArrayList<String> dates(int number){
        ArrayList<String> array = new ArrayList<String>();
        return array;
    }
}
