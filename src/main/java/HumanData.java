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
        String fileRepository = "src/main/resources/";
        int rowNum = readRowNum();

        String[] names = new String[rowNum],
                patronymics = new String[rowNum],
                dates = new String[rowNum],
                ages = new String[rowNum],
                index = new String[rowNum],
                countries = new String[rowNum],
                houses = new String[rowNum],
                flats = new String[rowNum],
                sexes = new String[rowNum],
                hometowns = new String[rowNum],
                regions = new String[rowNum],
                cities = new String[rowNum],
                streets = new String[rowNum],
                surnames = new String[rowNum];

        int daysInYear = 365;
        int hundredYearsInDays = 100 * daysInYear;

        ArrayList<String> femaleNamesList = getArrayFromFile(fileRepository + "FemaleNames.txt");
        ArrayList<String> maleNamesList = getArrayFromFile(fileRepository + "MaleNames.txt");
        ArrayList<String> surnamesList = getArrayFromFile(fileRepository + "Surnames.txt");
        ArrayList<String> femalePatronymicList = getArrayFromFile(fileRepository + "FemalePatronymic.txt");
        ArrayList<String> malePatronymicList = getArrayFromFile(fileRepository + "MalePatronymic.txt");
        ArrayList<String> streetsList = getArrayFromFile(fileRepository + "Streets.txt");
        ArrayList<String> citiesList = getArrayFromFile(fileRepository + "Cities.txt");
        ArrayList<String> regionsList = getArrayFromFile(fileRepository + "Regions.txt");
        ArrayList<String> hometownsList = getArrayFromFile(fileRepository + "Cities.txt");

        for (int i = 0; i < rowNum; i++){
            int randomCountOfDays = randomNumber(hundredYearsInDays);
            dates[i] = getData(randomCountOfDays);
            ages[i] = getAge(randomCountOfDays, daysInYear);
            index[i] = Integer.toString(100000 + randomNumber(899999));
            countries[i] = "Россия";
            houses[i] = Integer.toString(randomNumber(300));
            flats[i] = Integer.toString(randomNumber(300));
            if(randomNumber(2) == 1){
                sexes[i] = "MУЖ";
                names[i] = maleNamesList.get(randomNumber(maleNamesList.size() - 1));
                surnames[i] = surnamesList.get(randomNumber(surnamesList.size() - 1));
                patronymics[i] = malePatronymicList.get(randomNumber(malePatronymicList.size() - 1));
            } else {
                sexes[i] = "ЖЕН";
                names[i] = femaleNamesList.get(randomNumber(femaleNamesList.size() - 1));
                surnames[i] = surnamesList.get(randomNumber(surnamesList.size() - 1)) + "а";
                patronymics[i] = femalePatronymicList.get(randomNumber(femalePatronymicList.size()));
            }
            hometowns = getRandomArray(hometownsList, rowNum);
            regions = getRandomArray(regionsList, rowNum);
            cities = getRandomArray(citiesList, rowNum);
            streets = getRandomArray(streetsList, rowNum);
        }

        HSSFWorkbook workBook = new HSSFWorkbook();
        HSSFSheet sheet = workBook.createSheet("Тестовые данные");
        int i =0;
        while (i <= rowNum) {
            sheet.createRow(i);
            i++;
        }

        String[] columns  = {
                "Имя",
                "Фамилия",
                "Отчество",
                "Возраст",
                "Пол",
                "Дата рождения",
                "Место рождения",
                "Индекс",
                "Страна",
                "Область",
                "Город",
                "Улица",
                "Дом",
                "Квартира"
        };

        String[][] allData = {
                names,
                surnames,
                patronymics,
                ages,
                sexes,
                dates,
                hometowns,
                index,
                countries,
                regions,
                cities,
                streets,
                houses,
                flats
        };

        for (int columnIndex = 0; columnIndex < columns.length; columnIndex++){
            fillTheColumn(sheet, rowNum, columnIndex, allData[columnIndex], columns[columnIndex]);
        }

        File file = new File(fileRepository + "data.xls");
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

    private static void fillTheColumn(HSSFSheet sheet, int rowNum, int columnNum, String[] array, String columnName){
        String cellValue;
        Cell cell;
        Row row;
        row = sheet.getRow(0);
        cell = row.createCell(columnNum, CellType.STRING);
        cell.setCellValue(columnName);
        int i = 1;
        while (i <= rowNum) {
            cellValue = array[i-1];
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

    private static String[] getRandomArray(ArrayList<String> originArray, int newArrayLength){
        String[] newArray = new String[newArrayLength];
        for (int i = 0; i < newArrayLength; i ++) {
            newArray[i] = originArray.get(randomNumber(originArray.size() - 1));
        }
        return newArray;
    }
}
