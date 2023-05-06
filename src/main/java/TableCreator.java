import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Scanner;

public class TableCreator {
    public static void main(String[] args) throws IOException {

        String tableName;
        String fileLocation;
        Scanner sc =new Scanner(System.in);
        System.out.print("Entrer le chemin du fichier excel");
        fileLocation = sc.nextLine();
        System.out.print("le chemin est: "+fileLocation );
        FileInputStream file = new FileInputStream(fileLocation);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        System.out.println("ouverture de fichier reusite et localisation sur la feuil: "+sheet.getSheetName());
        tableName= sheet.getSheetName();
        int nbrPropriete;
        System.out.println("le nompbre totale de lighe est de :"+sheet.getLastRowNum());
        nbrPropriete=sheet.getLastRowNum();
        /*for(Row row:sheet){
              nbrPropriete++;
        }*/
        /*Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            System.out.println(nbrPropriete);
            nbrPropriete++;
        }*/
        System.out.println("le nembre total des propriété est: "+nbrPropriete+" propriétés");
        String tablePropNameList[]=new String[nbrPropriete];
        for(int i=0; i<nbrPropriete;i++){
            tablePropNameList[i]=sheet.getRow(i).getCell(0).getStringCellValue();
        }
        for(String pro:tablePropNameList){
            System.out.println(pro);
        }
        String tablePropTypeList[]= new String[nbrPropriete];
        for(int i=0; i<nbrPropriete;i++){
            tablePropTypeList[i]=sheet.getRow(i).getCell(1).getStringCellValue();
        }
        for(String type:tablePropTypeList){
            System.out.println(type);
        }
        String columnsOfStatment="";
        for(int i=0; i<nbrPropriete;i++){
            columnsOfStatment+=tablePropNameList[i]+" "+tablePropTypeList[i]+",";
        }
        String creationStatement;
        creationStatement="CREATE TABLE "+tableName+"("
                +"id_"+tableName+" serial,"
                +columnsOfStatment
                +" CONSTRAINT "+tableName+"_pkey PRIMARY KEY (id_"+tableName+")"
                +");";
        System.out.println(creationStatement);
       int nbrColumns=0;
        for(Row row:sheet){
            for(Cell cell:row){
                nbrColumns++;
            }
            nbrColumns-=2;
            break;
        }
        String columnsOfInsertionStmt="";
        String comma=",";
        for(int i=2;i<nbrColumns+2;i++){
            columnsOfInsertionStmt+="INSERT INTO "+tableName+" (";
            for(int j=0; j<nbrPropriete;j++){
                if (j==nbrPropriete-1) comma=")VALUES\n";
                columnsOfInsertionStmt+=tablePropNameList[j]+comma;
            }
            String valuesInStmt="(";
            for(int j=0; j<nbrPropriete;j++){
                Cell cell=(sheet.getRow(j).getCell(i));
                switch (cell.getCellType()){
                    case STRING:valuesInStmt+="'"+cell.getStringCellValue()+"'";
                        break;
                    case NUMERIC:valuesInStmt+=cell.getNumericCellValue();
                        break;
                }
                if(j<nbrPropriete-1){
                    valuesInStmt+=",";
                }else{
                    valuesInStmt+=");\n";
                }
            }
            columnsOfInsertionStmt+=valuesInStmt;

            comma=",";
        }
        //System.out.println(columnsOfInsertionStmt);
        creationStatement+=" "+columnsOfInsertionStmt;
         final String url = "jdbc:postgresql://localhost:5432/OPDATABASE";
         final String user = "admin";
         final String password = "123456789";
        Connection conn = null;


        try {
            conn = DriverManager.getConnection(url,user,password);
            System.out.println("Connected to the PostgreSQL server successfully.");
            PreparedStatement stmt = conn.prepareStatement(creationStatement);
            ResultSet rs= stmt.executeQuery();
            conn.commit();

        } catch (SQLException e) {
            System.out.println(e);
        }
    }
}
