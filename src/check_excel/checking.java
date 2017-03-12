/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package check_excel;

import java.io.FileInputStream;
import java.util.ArrayList;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Класс, в котором выполняются все основные действия: выдача уведомления об ошибке, чтение данных из файла .xls,
 * выдача финального уведомления о валидности документа
 * @author Alexandr
 */
public class checking {
    /**
     * единственная "видимая" переменная, через которую программа обращается к файлу. 
     * Является строкой, которую программа получает из графической формы
     */
    public String wave_to_doc;
    private ArrayList<String> First_Name = new ArrayList<String>();
    private ArrayList<String> Last_Name = new ArrayList<String>();
    private ArrayList<String> Email_Adress = new ArrayList<String>();
    private ArrayList<String> Password = new ArrayList<String>();
    private ArrayList<String> Secondary_Email = new ArrayList<String>();
    private ArrayList<String> Mobile_Phone = new ArrayList<String>();
    private ArrayList<String> Department = new ArrayList<String>();
    private Boolean final_result=true;
    /**
     * 
     * @param wave_to_doc1 - строка адреса расположения файла 
     */
    private void reading(String wave_to_doc1){
       
            try{
                HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(wave_to_doc1));
                HSSFSheet myExcelSheet = myExcelBook.getSheet("Data");
                ArrayList<HSSFRow> rows = new ArrayList<HSSFRow>();
                int  i_row_exc = 1;
                while(true){
                    rows.add(myExcelSheet.getRow(i_row_exc));
                    try{
                    String s = rows.get(rows.size()-1).getCell(0).getStringCellValue();
                    }
                    catch(Exception e){
                        break;
                    }
                    First_Name.add(rows.get(rows.size()-1).getCell(7).getStringCellValue());
                    Last_Name.add(rows.get(rows.size()-1).getCell(8).getStringCellValue());
                    Email_Adress.add(rows.get(rows.size()-1).getCell(2).getStringCellValue());
                    Password.add(rows.get(rows.size()-1).getCell(3).getStringCellValue());
                    Secondary_Email.add(rows.get(rows.size()-1).getCell(4).getStringCellValue());
                    Mobile_Phone.add(rows.get(rows.size()-1).getCell(5).getStringCellValue());
                    Department.add(rows.get(rows.size()-1).getCell(6).getStringCellValue());
                    i_row_exc++;
                }
            }
            catch (Exception e){
                JOptionPane.showMessageDialog(null, e);
            }
     
    }
    /**
     * функция, в которой проводится проверка значений на соответствия требованиям
     */
    private void checking(){
        int i = 0;
        Boolean gert;
        if (First_Name.size() == Last_Name.size() && First_Name.size() == Email_Adress.size() && First_Name.size() == Password.size() && First_Name.size() == Secondary_Email.size() && First_Name.size() ==Mobile_Phone.size() && First_Name.size() == Department.size())
            gert = true;
        else
            gert = false;
        if (gert)
            while(i < First_Name.size()){
                String first_name = First_Name.get(i);
                String second_name = Last_Name.get(i); 
                first_name = first_name.toLowerCase();
                second_name = second_name.toLowerCase();
                first_name  = first_name.trim();
                first_name = first_name.substring(0, 1);
                second_name = second_name.trim();
                //формирование строки для сравнения с почтой
                String result;
                int a = i+2;
                result = (first_name+"."+second_name+"@student.khai.edu");
                if (!result.equals(Email_Adress.get(i))){
                    JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Email неверен </i>" + "Строка данных: " + a);
                    final_result = false;
                }
                String passw = Password.get(i);
                if (passw.length()!=8){
                    JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Количество символов в пароле неверно </i>" + "Строка данных: " + a);
                final_result = false;
                }
                String number = Mobile_Phone.get(i);
                if (number.length()!= 13){
                    JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Количество символов в номере телефона неверно </i>" + "Строка данных: " + a);
                    final_result = false;
                    }
                i++;
            }
        if (!gert){
            JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Заполнены не все строки </i>");
            
        }
    }
    /**
     * Функция с модификатором public, которая включает в себя вызов функций  checking и reading с 
     * модификаторами private, выдает финальные уведомления о верности/неверности составления файла .xls
     */
    public void messages(){
        reading(wave_to_doc);
        checking();
        if (final_result){
            JOptionPane.showMessageDialog(null,"<html><h2Кажется, все верно</h2><i>Но проверьте еще разок</i>");
        }
        if(!final_result){
            JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Исправьте ошибки в указанных строках и повторите попытку </i>");
        }
    }
}
