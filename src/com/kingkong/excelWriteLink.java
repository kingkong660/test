package demo;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Hyperlink;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;

public class excelWriteLink {
	public static void main(String[] str) {
		try {
			System.out.println("开始");
			Scanner sc = new Scanner(System.in);
			// 文件路径
			String strFileN = "";
			// 开始行
			String strRow = "";
			// 案例列
			String strBegingCol = "";
			System.out.println("输入Excel绝对路径: 例(F:/测试案例.xls 注:Excel文件类型只支持xls格式)");
			strFileN = sc.nextLine();  //读取字符串型输入
			System.out.println("输入案例开始读取行: 例(2 从第2行开始读取 注:0代表第一行)");
			strRow = sc.nextLine();  //读取字符串型输入
			System.out.println("输入案例列索引: 例(0 从第0列开始读取 注:0代表第一列)");
			strBegingCol = sc.nextLine();  //读取字符串型输入
//			strFileN="E://1.xls";
//			strRow="3";
//			strBegingCol="0";

			excelWriteLink cExcel = new excelWriteLink();

			File cFile = new File(strFileN);
			if (!cFile.exists()) {
				System.out.println("没有找到文件!(" + strFileN + ")");
				return;
			}
			if (!isNumber(strRow)) {
				System.out.println("开始行不是有效数字!(" + strRow + ")");
				return;
			}
			if (!isNumber(strBegingCol)) {
				System.out.println("开始行不是有效数字!(" + strBegingCol + ")");
				return;
			}
			cExcel.m_nReadCol = Integer.parseInt(strBegingCol);
			cExcel.m_nBeginRow = Integer.parseInt(strRow);
			cExcel.initExcel2003(strFileN);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 记录案例数据
	 */
	public Map<String, String> m_MapOriginalCase = new HashMap<String, String>();
	/**
	 * 记录结果输出案例信息
	 */
	public Map<String, String> m_MapResultCase = new HashMap<String, String>();
	/**
	 * 案例索引列
	 */
	public int m_nReadCol = 0;
	/**
	 * 目标输出列
	 */
	public int m_nTargetCol = 0;
	/**
	 * 开始行
	 */
	public int m_nBeginRow = 0;

	// 按照2003excel版本读取2003版excel文件，复制excel样式
	private void initExcel2003(String strPath) throws Exception {
		File cRFile = new File(strPath);
		FileInputStream cRFileInput = new FileInputStream(cRFile);
		HSSFWorkbook cRwork = new HSSFWorkbook(cRFileInput);
//		System.out.println("sheet->" + cRwork.getNumberOfSheets());
		HSSFSheet cRsheet = cRwork.getSheetAt(0);
		int nSheetCount = cRwork.getNumberOfSheets();
		int nRowMax = cRsheet.getLastRowNum();
		if (m_nBeginRow > nRowMax) {
			System.out.println("数据读取行大于Excel里大行!("+m_nBeginRow+" "+nRowMax+")");
			return;
		}
		// 设置excel单元格样式
		for (int nR = m_nBeginRow; nR <= nRowMax; nR++) {
			HSSFRow cReRow = cRsheet.getRow(nR);
			if(cReRow ==null) continue;
			int nColMax = cReRow.getLastCellNum();
			if (nColMax >= m_nReadCol) {
				HSSFCell cRCel = cReRow.getCell(m_nReadCol);
				if (cRCel != null) {
					// 复制表格内容
					cRCel.setCellType(HSSFCell.CELL_TYPE_STRING);
					String strV = cRCel.getStringCellValue();
					if (strV.length() > 0) {
						m_MapOriginalCase.put(strV, nR + "");
					}
				}
			}
		}
		// 循环Sheet页
		for (int nSheetIndex = 1; nSheetIndex < nSheetCount; nSheetIndex++) {
			cRsheet = cRwork.getSheetAt(nSheetIndex);
			nRowMax = cRsheet.getLastRowNum();
			// 循环行
			for (int nR = 0; nR <= nRowMax; nR++) {
				HSSFRow cReRow = cRsheet.getRow(nR);
				if (cReRow != null) {
					int nColMax = cReRow.getLastCellNum();
					// 循环列
					for (int nCol = 0; nCol < nColMax; nCol++) {
						HSSFCell cRCel = cReRow.getCell(nCol);
						if (cRCel != null) {
							cRCel.setCellType(HSSFCell.CELL_TYPE_STRING);
							String strV = cRCel.getStringCellValue();
							if (strV.trim().length() > 0) {
								if (m_MapOriginalCase.containsKey(strV)) {
									// System.out.println(strV + " " + nR + " "
									// + nCol);
									m_MapResultCase.put(strV, m_MapOriginalCase.get(strV) + "_" + nSheetIndex + "_" + nR + "_" + nCol);
									m_MapOriginalCase.remove(strV);
								}
							}
						}
					}
				}
			}
		}
		// 获取案例对象
		HSSFSheet cSheetIndex1 = cRwork.getSheetAt(0);
		for (String key : m_MapResultCase.keySet()) {
//			System.out.println(key + " " + m_MapResultCase.get(key));
			String strParms[] = m_MapResultCase.get(key).split("_");
			HSSFRow cReRow = cSheetIndex1.getRow(Integer.parseInt(strParms[0]));
			if (cReRow != null) {
				HSSFCell cRCel = cReRow.getCell(m_nTargetCol);
				if (cRCel != null) {
					String strSheetName = cRwork.getSheetName(Integer.parseInt(strParms[1]));
//					System.out.println("#" + strSheetName + "!A" + (Integer.parseInt(strParms[2]) + 1));
					Hyperlink hyperlink = new HSSFHyperlink(Hyperlink.LINK_DOCUMENT);
					hyperlink.setAddress("#" + strSheetName + "!"+getExcelColumnLabel(Integer.parseInt(strParms[3]))+(Integer.parseInt(strParms[2]) + 1));
					cRCel.setHyperlink(hyperlink);

					HSSFCellStyle linkStyle = cRwork.createCellStyle();
					HSSFFont cellFont = cRwork.createFont();
					cellFont.setUnderline((byte) 1);
					cellFont.setColor(HSSFColor.BLUE.index);
					linkStyle.setFont(cellFont);
					cRCel.setCellStyle(linkStyle);

				}
			}
		}

		FileOutputStream outStream = new FileOutputStream(strPath);
		System.out.println(strPath);
		cRwork.write(outStream);
		outStream.flush();
		outStream.close();

		System.out.println("案例数共:"+(m_MapOriginalCase.size()+m_MapResultCase.size()));
		System.out.println("匹配成功:"+(m_MapResultCase.size()));
		System.out.println("未匹配成功:"+m_MapOriginalCase.size());
		if(m_MapOriginalCase.size()>0){
			String strMsg="";
			for (String key : m_MapOriginalCase.keySet()) {
				strMsg += (key+" ");
			}
			System.out.println("未找到案例:["+strMsg+"]");
		}
		System.out.println("执行结束!");
	}

	public String getExcelColumnLabel(int num) {
		String temp = "";
		double i = Math.floor(Math.log(25.0 * (num) / 26.0 + 1) / Math.log(26)) + 1;
		if (i > 1) {
			double sub = num - 26 * (Math.pow(26, i - 1) - 1) / 25;
			for (double j = i; j > 0; j--) {
				temp = temp + (char) (sub / Math.pow(26, j - 1) + 65);
				sub = sub % Math.pow(26, j - 1);
			}
		} else {
			temp = temp + (char) (num + 65);
		}
		return temp;
	}

	public static boolean isNumber(String strVal){
		try {
			Integer.parseInt(strVal);
		} catch (Exception e) {
			return false;
		}
		return true;
	}
}
