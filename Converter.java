import java.io.File;
import java.util.ArrayList;

import gurobi.GRB;
import gurobi.GRBVar;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class Converter {
	/**
	 * Converts given Excel sheet into two dimensional ArrayList
	 * 
	 * @param address of file
	 * @param sheet   to be converted
	 * @return ArrayList with equivalent content
	 */
	public static ArrayList<ArrayList<Double>> excelToArrayList(String address, int sheet, int firstRow, int firstCol) {
		try {
			File f = new File(address);
			Workbook Wb = Workbook.getWorkbook(f);
			Sheet sh = Wb.getSheet(sheet);
			ArrayList<ArrayList<Double>> mathArray = new ArrayList<>();
			int row = sh.getRows();
			int col = sh.getColumns();
			for (int i = firstRow; i < row; i++) {
				ArrayList<Double> colArr = new ArrayList<>();
				for (int j = firstCol; j < col; j++) {
					Cell c = sh.getCell(j, i);
					colArr.add(Double.parseDouble(c.getContents().replace(',', '.')));
				}
				mathArray.add(colArr);
			}
			return mathArray;
		} catch (Exception e) {
			System.out.println("Fehler: " + e.getMessage());
			return null;
		}
	}
	
	/**
	 * Converts two dimensional ArrayList into two dimensional array
	 * 
	 * @param mathArray ArrayList to be converted
	 * @return Array with equivalent content
	 */
	public static double[][] arrayListToArray(ArrayList<ArrayList<Double>> mathArray) {
		int rows = mathArray.size();
		int cols = mathArray.get(0).size();
		double[][] res = new double[rows][cols];
		for (int i = 0; i < rows; i++) {
			for (int j = 0; j < cols; j++) {
				res[i][j] = mathArray.get(i).get(j);
			}
		}
		return res;
	}
	
	/**
	 * Converts two dimensional ArrayList into two dimensional array
	 * 
	 * @param mathArray               ArrayList to be converted
	 * @param rows                    number of of the result (should be larger than
	 *                                numberOfGroups)
	 * @param cols                    number of columns of the result (should be
	 *                                larger than each element of
	 *                                numberOfStudentsInGroup)
	 * @param numberOfGroups          length of superlist
	 * @param numberOfStudentsInGroup length of each sublist
	 * @return Array with equivalent content
	 */
	public static double[][] arrayListToArray(ArrayList<ArrayList<Double>> mathArray, int rows, int cols,
			int numberOfGroups, int[] numberOfStudentsInGroup) {
		double[][] res = new double[rows][cols];
		for (int j = 0; j < numberOfGroups; j++) {
			for (int i = 0; i < numberOfStudentsInGroup[j]; i++) {
				res[(int) Math.round(mathArray.get(j).get(i))][j] = 1;
			}
		}
		return res;
	}
	
	/**
	 * Converts from continuous [-1,1] preferences to discrete {0,1} preferences
	 * 
	 * @param res array to be normalized
	 * @return array with binary preferences
	 */
	public static double[][] binaryPreferences(double[][] inp) {
		int rows = inp.length;
		int cols = inp[0].length;
		double[][] res = new double[rows][cols];

		for (int i = 0; i < rows; i++) {
			double maxInRow = 0;
			int indexMaxInRow = -1;
			for (int j = 0; j < cols; j++) {
				if (inp[i][j] > maxInRow) {
					maxInRow = inp[i][j];
					indexMaxInRow = j;
				}
			}
			for (int j = 0; j < cols; j++)
				res[i][j] = (j == indexMaxInRow) ? 1 : 0;
		}
		return res;
	}
	
	/**
	 * Converts from inter-social x_ss variables to social-group x_sg variables
	 * 
	 * @param x inter-social x_ss variables
	 * @param G max. number of groups
	 * @return social-group x_sg variables
	 */
	public static double[][] convertFromXtoXsg(GRBVar[][] x, int G) {
		try {
			int S = x.length;
			double[][] y = new double[S][S];
			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					y[i][j] = x[i][j].get(GRB.DoubleAttr.X);
				}
			}

			int[] numberOfStudentsInGroup = new int[G];
			int numberOfGroups = 0;
			ArrayList<ArrayList<Double>> res = new ArrayList<>();
			for (int i = 0; i < S; i++) {
				if (y[i][i] < -1 / 2)
					continue;
				numberOfGroups++;
				ArrayList<Double> rowArr = new ArrayList<>();
				rowArr.add(i * 1.0);
				numberOfStudentsInGroup[numberOfGroups - 1]++;
				for (int j = i + 1; j < S; j++) {
					if (y[i][j] > 1 / 2) {
						rowArr.add(j * 1.0);
						y[j][j] = -1;
						numberOfStudentsInGroup[numberOfGroups - 1]++;
					}
				}
				res.add(rowArr);
			}
			return arrayListToArray(res, S, G, numberOfGroups, numberOfStudentsInGroup);
		} catch (Exception e) {
			System.out.println("Fehler: " + e.getMessage());
			return null;
		}
	}
	
	/**
	 * Normalizes non-zero rows to L1-norm = 1
	 * 
	 * @param res array to be normalized
	 * @return array with L1-normalized rows
	 */
	public static double[][] normalizePreferences(double[][] res) {
		int rows = res.length;
		int cols = res[0].length;

		for (int i = 0; i < rows; i++) {
			double rowSum = 0;
			for (int j = 0; j < cols; j++) {
				rowSum += Math.abs(res[i][j]);
			}
			if (rowSum != 0) {
				for (int j = 0; j < cols; j++) {
					res[i][j] = res[i][j] / rowSum;
				}
			}
		}
		return res;
	}
}
