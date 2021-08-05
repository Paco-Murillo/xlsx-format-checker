import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Scanner;
import java.util.regex.Pattern;

public class XLSXFormatChecker {

    private static final String[] xlsxColumns = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
    private static final Pattern emailPattern = Pattern.compile("^[a-zA-Z0-9]+(?:[-_.][a-zA-Z0-9]+)*@[a-z0-9]+(?:\\.+[a-z0-9\\-]{2,})+$");
    private static final Pattern domainPattern = Pattern.compile("^[a-zA-Z0-9]+(?:[-_.][a-zA-Z0-9]+)*@[a-z0-9]+\\.con\\.*.*");
    private static final Pattern namePattern = Pattern.compile("^[A-Za-zÀ-ÖØ-öø-ÿ]+\\.*(?:\\s(?:[A-Za-zÀ-ÖØ-öø-ÿ]*)*\\.*)*$");
    private static final Pattern numberPattern = Pattern.compile("^[0-9]+$");
    
    private enum CellType{
        EMAIL,
        NUMBER,
        NAME
    }

    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("altas.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);

        CellType[] cellTypeChecker = new CellType[row.getLastCellNum()];
        CellType[] cellTypes = CellType.values();
        Scanner scanner = new Scanner(System.in);

        System.out.println("Para cada una de las columnas seleccionar el tipo de información a checar por sintaxis\n" +
                "1 - Email\n2 - Numero de telefono\n3 - Nombre\n");

        XSSFCell c;
        for (int i = 0; i < row.getLastCellNum(); i++) {
            c = row.getCell(i);
            System.out.print(c +": ");
            cellTypeChecker[i] = cellTypes[scanner.nextInt() - 1];
        }

        String stringValueOfCell;
        boolean email = true, number = true, name = true;
        for (int rowIndex = 1; rowIndex < sheet.getLastRowNum(); rowIndex++) {
            row = sheet.getRow(rowIndex);
            for (int columnIndex = 0; columnIndex < row.getLastCellNum(); columnIndex++) {
                c = row.getCell(columnIndex);
                stringValueOfCell = c.toString();

                switch (cellTypeChecker[columnIndex]){
                    case EMAIL:
                        if (!emailPattern.matcher(stringValueOfCell).matches() || domainPattern.matcher(stringValueOfCell).matches()){
                            printCheckCell(rowIndex,columnIndex);
                            if(email) email = false;
                        }
                        break;
                    case NAME:
                        if (!namePattern.matcher(stringValueOfCell).matches()){
                            char[] chars = stringValueOfCell.toCharArray();
                            System.out.println(Arrays.toString(chars));
                            printCheckCell(rowIndex,columnIndex);
                            if(name) name = false;
                        }
                        break;
                    case NUMBER:
                        if (!numberPattern.matcher(stringValueOfCell).matches()){
                            printCheckCell(rowIndex,columnIndex);
                            if(number) number = false;
                        }
                        break;
                    default:
                        throw new Error("Error en Enum CellTypes");
                }
            }
        }
        fis.close();
    }

    private static void printCheckCell(int row, int cell){
        row++;
        System.out.println("Checar celda: " + intToColumnName(cell) + row);
    }

    private static String intToColumnName(int cell){
        return cell>=xlsxColumns.length? intToColumnName((cell/xlsxColumns.length)-1)+xlsxColumns[cell%xlsxColumns.length]: xlsxColumns[cell%xlsxColumns.length];
    }
}
