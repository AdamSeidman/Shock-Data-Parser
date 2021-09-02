package seidman;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Util {

	public static File getWorkbook() {
		JFileChooser chooser = new JFileChooser();
		chooser.setCurrentDirectory(new File("./"));
		chooser.addChoosableFileFilter(new FileFilter() {
			public String getDescription() {
				return "Excel Documents (*.xlsx)";
			}

			public boolean accept(File f) {
				if (f.isDirectory()) {
					return true;
				} else {
					return f.getName().toLowerCase().endsWith(".xlsx");
				}
			}
		});
		chooser.setAcceptAllFileFilterUsed(false);
		chooser.showDialog(null, "Accept");
		File file = chooser.getSelectedFile();
		if (file == null || !file.exists() || file.isDirectory() || !file.getAbsolutePath().endsWith(".xlsx")) {
			showMessage("There was an error with your file.", true);
			System.exit(1);
		}
		return file;
	}

	public static void showMessage(String message, boolean isError) {
		JOptionPane.showConfirmDialog(null, new JLabel(message, SwingConstants.CENTER), isError ? "Error" : "Message",
				JOptionPane.DEFAULT_OPTION, isError ?JOptionPane.ERROR_MESSAGE : JOptionPane.PLAIN_MESSAGE, null);
	}

	public static void setLook() {
		try {
			UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
		} catch (Exception ignored) {
		}
	}

	public static void startParsingThread() {
		new Thread() {
			public void run() {
				JLabel label = new JLabel("Parsing", SwingConstants.CENTER);
				Thread t = new Thread() {
					public void run () {
						int n = 0;
						while (true) {
							try {
								Thread.sleep(250);
							} catch (InterruptedException ignored) {
							}
							n = (n + 1) % 5;
							label.setText((label.getText() + ".....").substring(0,7 + n));
						}
					}
				};
				t.start();
				JOptionPane.showOptionDialog(null, label, "Message", JOptionPane.DEFAULT_OPTION, 
						JOptionPane.PLAIN_MESSAGE, null, new Object[] {}, null);
				t.interrupt();
			}
		}.start();
	}
	
	public static double getCoefficient(XSSFSheet sheet) {
		double xSum = 0.0, ySum = 0.0, xSqSum = 0.0, xySum = 0.0, n = 0.0;
		boolean gotTitles = false;

		for (Row row : sheet) {
			if (!gotTitles) {
				gotTitles = true;
			} else {
				n++;
				double x = row.getCell(0).getNumericCellValue();
				x *= 0.0254;
				double y = row.getCell(1).getNumericCellValue();
				y *= 4.44822;
				xSum += x;
				ySum += y;
				xSqSum += (x * x);
				xySum += (x * y);
			}
		}

		double result = (n * xySum);
		result -= (xSum * ySum);
		double denom = (xSqSum * n);
		denom -= Math.pow(xSum, 2.0);
		result = (result / denom);

		return result;
	}

	public static void writeResults(XSSFWorkbook wb, FileOutputStream fos) throws IOException{
		XSSFSheet results = wb.createSheet("Results");
		Row row_ = results.createRow(0);
		row_.createCell(0).setCellValue("Sheet Name");
		row_.createCell(1).setCellValue("Damping Coefficient");
		row_.createCell(2).setCellValue("Damping Ratio (60 lb spring)");
		int n = 1;
		for(String name : points.keySet()) {
			Row row = results.createRow(n);
			row.createCell(0).setCellValue(name);
			row.createCell(1).setCellValue(points.get(name));
			double ratio = (points.get(name) / 1618.933547);
			row.createCell(2).setCellValue(ratio);
			
			n++;
		}
		wb.write(fos);
		try {
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static HashMap<String, Double> points = new HashMap<String, Double>();

	public static void deleteResultsSheet(XSSFWorkbook wb, FileOutputStream fos) {
		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			if (wb.getSheetAt(i).getSheetName().equals("Results")) {
				wb.removeSheetAt(i);
				try {
					wb.write(fos);
				} catch (IOException e) {
					e.printStackTrace();
					try {
						fos.close();
					} catch (IOException e1) {
						e1.printStackTrace();
					}
					showMessage("Could not delete old results.", true);
					System.exit(1);
				}
				return;
			}
		}
	}
	
	public static void main(String[] args) {
		setLook();
		XSSFWorkbook workbook = null;
		File file = null;
		FileOutputStream fos = null;
		try {
			file = getWorkbook();
			startParsingThread();
			workbook = new XSSFWorkbook(new FileInputStream(file));
			fos = new FileOutputStream(file);
		} catch (IOException e) {
			e.printStackTrace();
			showMessage("Could not open workbook.", true);
			System.exit(1);
		}
		deleteResultsSheet(workbook, fos);
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			XSSFSheet sheet = workbook.getSheetAt(i);
			if (!sheet.getSheetName().trim().equalsIgnoreCase("master")) {
				points.put(sheet.getSheetName(), getCoefficient(sheet));
			}
		}
		try {
			writeResults(workbook, fos);
			JOptionPane.getRootFrame().dispose();
			showMessage("Complete!", false);
		} catch (IOException e) {
			JOptionPane.getRootFrame().dispose();
			e.printStackTrace();
			showMessage("Could not write results.", true);
		}
	}

}
