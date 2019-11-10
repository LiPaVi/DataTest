import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.text.SimpleDateFormat;
import java.util.*;

public class HumanData {
    public static void main(String[] args) throws Exception{
        int rowNum = readRowNum();

        ArrayList<String> names = new ArrayList<String>(),
                patronymics = new ArrayList<String>(),
                dates = new ArrayList<String>(),
                ages = new ArrayList<String>(),
                index = new ArrayList<String>(),
                countries = new ArrayList<String>(),
                houses = new ArrayList<String>(),
                flats = new ArrayList<String>(),
                sexes = new ArrayList<String>();
        int daysInYear = 365;
        int hundredYearsInDays = 100 * daysInYear;

        ArrayList<String> femaleNames = getArrayFromFile("src/main/resources/FemaleNames.txt");
        ArrayList<String> maleNames = getArrayFromFile("src/main/resources/MaleNames.txt");
        ArrayList<String> surnames = getArrayFromFile("src/main/resources/Surnames.txt");
        ArrayList<String> femalePatronymic = getArrayFromFile("src/main/resources/FemalePatronymic.txt");
        ArrayList<String> malePatronymic = getArrayFromFile("src/main/resources/MalePatronymic.txt");
        ArrayList<String> streets = getArrayFromFile("src/main/resources/Streets.txt");
        ArrayList<String> cities = getArrayFromFile("src/main/resources/Cities.txt");
        ArrayList<String> regions = getArrayFromFile("src/main/resources/Regions.txt");

        for (int i = 0; i < rowNum; i++){
            int randomCountOfDays = randomNumber(hundredYearsInDays);
            dates.add(i, getData(randomCountOfDays));
            ages.add(i, getAge(randomCountOfDays, daysInYear));
            index.add(i, Integer.toString(100000 + randomNumber(899999)));
            countries.add(i, "Россия");
            houses.add(i, Integer.toString(randomNumber(300)));
            flats.add(i, Integer.toString(randomNumber(300)));
            if(randomNumber(2) == 1){
                sexes.add(i, "MУЖ");
                names.add(i, maleNames.get(randomNumber(maleNames.size())));
                surnames.add(i, surnames.get(randomNumber(surnames.size())));
                patronymics.add(i, malePatronymic.get(randomNumber(malePatronymic.size())));
            } else {
                sexes.add(i, "ЖЕН");
                names.add(i, femaleNames.get(randomNumber(femaleNames.size())));
                surnames.add(i, surnames.get(randomNumber(surnames.size())) + "а");
                patronymics.add(i, femalePatronymic.get(randomNumber(femalePatronymic.size())));
            }
        }

        HSSFWorkbook workBook = new HSSFWorkbook();
        HSSFSheet sheet = workBook.createSheet("Тестовые данные");
        int i =0;
        while (i <= rowNum) {
            sheet.createRow(i);
            i++;
        }

        fillTheColumn(sheet, rowNum, 0, names, "Имя");
        fillTheColumn(sheet, rowNum, 1, surnames, "Фамилия");
        fillTheColumn(sheet, rowNum, 2, patronymics, "Отчество");
        fillTheColumn(sheet, rowNum, 3, ages, "Возраст");
        fillTheColumn(sheet, rowNum, 4, sexes, "Пол");
        fillTheColumn(sheet, rowNum, 5, dates, "Дата рождения");
        fillTheColumnWithRandom(sheet, rowNum, 6, cities, "Место рождения");
        fillTheColumn(sheet, rowNum, 7, index, "Индекс");
        fillTheColumn(sheet, rowNum, 8, countries, "Страна");
        fillTheColumnWithRandom(sheet, rowNum, 9, regions, "Область");
        fillTheColumnWithRandom(sheet, rowNum, 10, cities, "Город");
        fillTheColumnWithRandom(sheet, rowNum, 11, streets, "Улица");
        fillTheColumn(sheet, rowNum, 12, houses, "Дом");
        fillTheColumn(sheet, rowNum, 13, flats, "Квартира");
        File file = new File("src/main/resources/data.xls");
        FileOutputStream outFile = new FileOutputStream(file);
        workBook.write(outFile);
        System.out.println("Файл создан. Путь: " + file.getAbsolutePath());
        workBook.close();
    }

    private static int readRowNum(){
        Scanner in = new Scanner(System.in);
        int rowNum = 0;
        System.out.print("Введите число от 1 до 30: ");
        try {
            rowNum = in.nextInt();
        }
        catch (InputMismatchException e){
            System.out.println("Ошибка! Неверные входные данные!");
        }
        if (rowNum < 1 || rowNum > 30){
            System.out.println("Ваше число не соответствует условию." +
                    "\nБудет автоматически сгенерирован файл из одной строки.");
            rowNum = 1;
        }
        in.close();
        return rowNum;
    }

    private static int randomNumber(int max){
        return (int) (Math.random() * max);
    }

    private static void fillTheColumnWithRandom(HSSFSheet sheet, int rowNum, int columnNum, ArrayList<String> list, String columnName){
        String cellValue;
        Cell cell;
        Row row;
        row = sheet.getRow(0);
        cell = row.createCell(columnNum, CellType.STRING);
        cell.setCellValue(columnName);
        int i = 1;
        while (i <= rowNum) {
            cellValue = list.get(randomNumber(list.size()));
            row = sheet.getRow(i);
            cell = row.createCell(columnNum, CellType.STRING);
            cell.setCellValue(cellValue);
            i += 1;
        }
    }

    private static void fillTheColumn(HSSFSheet sheet, int rowNum, int columnNum, ArrayList<String> list, String columnName){
        String cellValue;
        Cell cell;
        Row row;
        row = sheet.getRow(0);
        cell = row.createCell(columnNum, CellType.STRING);
        cell.setCellValue(columnName);
        int i = 1;
        while (i <= rowNum) {
            cellValue = list.get(i-1);
            row = sheet.getRow(i);
            cell = row.createCell(columnNum, CellType.STRING);
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

    private static String getData(int days){
        Calendar date = new GregorianCalendar();
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        date.add(Calendar.DAY_OF_MONTH, - days);
        return dateFormat.format(date.getTime());
    }

    private static String getAge(int daysSinceBirth, int daysInYear){
        return Integer.toString(daysSinceBirth / daysInYear);
    }
}
