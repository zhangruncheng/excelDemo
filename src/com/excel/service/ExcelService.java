package com.excel.service;

import java.io.FileOutputStream;   
import java.io.IOException;   
import java.io.OutputStream;   
import java.sql.ResultSet;   
import java.sql.SQLException;   
import java.util.*;   
import javax.swing.JOptionPane;   
import org.apache.poi.hssf.usermodel.HSSFCell;   
import org.apache.poi.hssf.usermodel.HSSFFooter;   
import org.apache.poi.hssf.usermodel.HSSFHeader;   
import org.apache.poi.hssf.usermodel.HSSFRow;   
import org.apache.poi.hssf.usermodel.HSSFSheet;   
import org.apache.poi.hssf.usermodel.HSSFWorkbook; 
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;

public class ExcelService {
//  ��ͷ  
    public static final String[] tableHeader = {"id","����","����"};   
//  ����������  
    public static HSSFWorkbook demoWorkBook = new HSSFWorkbook();   
//  ������  
    public static HSSFSheet demoSheet = demoWorkBook.createSheet("�û���Ϣ");   
//  ��ͷ�ĵ�Ԫ�����Ŀ  
    public static final short cellNumber = (short)tableHeader.length;   
//  ���ݿ�������  
    public static final int columNumber = 3;   
    /** 
     * 34.* ������ͷ 35.* 
     *  
     * @return 36. 
     */   
    @SuppressWarnings("deprecation")  
    public static void createTableHeader()   
    {  
    	//���ñ�ͷ����sheet�еõ�
        HSSFHeader header = demoSheet.getHeader();   
        header.setCenter("�û���");   
        //����һ��
        HSSFRow headerRow = demoSheet.createRow((short) 0);   
        for(int i = 0;i < cellNumber;i++)   
        {   
        	//����һ����Ԫ��
            HSSFCell headerCell = headerRow.createCell((short) i);   
           // headerCell.setEncoding(HSSFCell.ENCODING_UTF_16); 
//            CellStyle cs = new CellStyle();
            //����cell��ֵ
            headerCell.setCellValue(tableHeader[i]);   
        }  
    }  
/** 
 * 50.* ������ 51.* 
 *  
 * @param cells 
 *            52.* 
 * @param rowIndex 
 */   
@SuppressWarnings("deprecation")  
public static void createTableRow(List<String> cells , short rowIndex)   
{   
//  ������rowIndex��  
    HSSFRow row = demoSheet.createRow((short) rowIndex);   
    for(short i = 0;i < cells.size();i++)   
    {   
//      ������i����Ԫ��  
        HSSFCell cell = row.createCell((short) i);   
        //cell.setEncoding(HSSFCell.ENCODING_UTF_16);   
        cell.setCellValue(cells.get(i));   
    }   
}   
/** 
 * 68.* ��������Excel�� 69.* 
 *  
 * @throws SQLException 
 *             70.* 71. 
 */   
public static void createExcelSheeet() throws Exception   
{   
    createTableHeader();   //--->����һ����ͷ��
    ResultSet rs = SheetDataSource.selectAllDataFromDB();   //--->�õ���������   
    int rowIndex = 1;   
    while(rs.next())   
    {   
        List<String> list = new ArrayList<String>();   
        for(int i = 1;i <= columNumber;i++)   
        {   
            list.add(rs.getString(i));   
        }   
        createTableRow(list,(short)rowIndex);   
        rowIndex++;   
    }   
}   
/** 
 * 89.* ������� 90.* 
 *  
 * @param sheet 
 *            91.* 
 * @param os 
 *            92.* 
 * @throws IOException 
 *             93. 
 */   
public void exportExcel(HSSFSheet sheet,OutputStream os) throws IOException   
{   
    sheet.setGridsPrinted(true);   
    HSSFFooter footer = sheet.getFooter();   
    footer.setRight("Page " + HSSFFooter.page() + " of " +   
            HSSFFooter.numPages());   
    demoWorkBook.write(os);   
}   
public static void main(String[] args) {   
    String fileName = "D:\\�û���Ϣ.xls";   
    FileOutputStream fos = null;   
    try {   
    	ExcelService pd = new ExcelService();   
    	ExcelService.createExcelSheeet();   
        fos = new FileOutputStream(fileName);   
        pd.exportExcel(demoSheet,fos);   
        JOptionPane.showMessageDialog(null, "����ѳɹ������� : "+fileName);   
    } catch (Exception e) {   
        JOptionPane.showMessageDialog(null, "��񵼳�����������Ϣ ��"+e+"\n����ԭ������Ǳ���Ѿ��򿪡�");   
        e.printStackTrace();   
    } finally {   
        try {   
            fos.close();   
        } catch (Exception e) {   
            e.printStackTrace();   
        }   
    }   
}   
}
