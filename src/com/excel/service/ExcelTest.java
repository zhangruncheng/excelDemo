package com.excel.service;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.Region;

public class ExcelTest {

	/**
	 * @param args
	 */
	@SuppressWarnings("deprecation") 
	public static void main(String[] args) throws IOException {

		try {
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("new   sheet");
			HSSFCellStyle style = wb.createCellStyle(); // ��ʽ����

			style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// ��ֱ
			style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// ˮƽ
			HSSFRow row = sheet.createRow((short) 0);
			HSSFRow row2 = sheet.createRow((short) 1);

			sheet.addMergedRegion(new Region(0, (short) 0, 2, (short) 0));
			HSSFCell ce = row.createCell((short) 0);
			// ce.setEncoding(HSSFCell.ENCODING_UTF_16);// ���Ĵ���
			ce.setCellValue("��Ŀ\\����"); // ���ĵ�һ�е�һ����ʾ������
			ce.setCellStyle(style); // ��ʽ������
			int num = 0;
			for (int i = 0; i < 9; i++) { // ѭ��9�Σ�ÿһ�ζ�Ҫ�絥Ԫ����ʾ
				// ������Ǹ���Ԫ��絽��һ��
				int celln = 0;
				int celle = 0;
				if (i == 0) {
					celln = 0;
					celle = 1;
				} else {
					celln = (i * 2);
					celle = (i * 2 + 1);
				}
				// ��Ԫ��ϲ�
				// �ĸ������ֱ��ǣ���ʼ�У���ʼ�У������У�������
				sheet.addMergedRegion(new Region(0, (short) (celln + 1), 0,
						(short) (celle + 1)));
				HSSFCell cell = row.createCell((short) (celln + 1));
				cell.setCellValue("merging" + i); // �絥Ԫ����ʾ������
				cell.setCellStyle(style); // ��ʽ
				// ���絥Ԫ����ʾ�����ݣ��磺�����У���һ�зֱ�����Ϊһ����һ�о�Ϊ���񣬡�������������
				HSSFCell cell1 = row2.createCell((short) celle);
				HSSFCell cell2 = row2.createCell((short) (celle + 1));
				// cell1.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell1.setCellValue("����");
				cell1.setCellStyle(style);
				// cell2.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell2.setCellValue("���");
				cell2.setCellStyle(style);
				num++;
			}

			// �ں�����Ϻϼưٷֱ�

			// �ϼ� �������ϣ���Ҫ��һ����Ԫ��
			sheet.addMergedRegion(new Region(0, (short) (2 * num + 1), 0,
					(short) (2 * num + 2)));
			HSSFCell cell = row.createCell((short) (2 * num + 1));
			// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("�ϼ�");
			cell.setCellStyle(style);
			HSSFCell cell1 = row2.createCell((short) (2 * num + 1));
			HSSFCell cell2 = row2.createCell((short) (2 * num + 2));
			// cell1.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell1.setCellValue("����");
			cell1.setCellStyle(style);
			// cell2.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell2.setCellValue("���");
			cell2.setCellStyle(style);

			// �ٷֱ� ͬ��
			sheet.addMergedRegion(new Region(0, (short) (2 * num + 3), 0,
					(short) (2 * num + 4)));
			HSSFCell cellb = row.createCell((short) (2 * num + 3));
			// cellb.setEncoding(HSSFCell.ENCODING_UTF_16);
			cellb.setCellValue("�ٷֱ�");
			cellb.setCellStyle(style);
			HSSFCell cellb1 = row2.createCell((short) (2 * num + 3));
			HSSFCell cellb2 = row2.createCell((short) (2 * num + 4));
			// cellb1.setEncoding(HSSFCell.ENCODING_UTF_16);
			cellb1.setCellValue("����");
			cellb1.setCellStyle(style);
			// cellb2.setEncoding(HSSFCell.ENCODING_UTF_16);
			cellb2.setCellValue("���");
			cellb2.setCellStyle(style);

			FileOutputStream fileOut = new FileOutputStream("workbook.xls");
			wb.write(fileOut);
			fileOut.close();
			System.out.print("OK");
		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

}
