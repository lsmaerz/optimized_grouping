import java.io.File;
import java.util.ArrayList;

import gurobi.GRB;
import gurobi.GRBEnv;
import gurobi.GRBException;
import gurobi.GRBLinExpr;
import gurobi.GRBModel;
import gurobi.GRBVar;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class StandardNaive {
	public static void main(String[] args) {
	}

	/**
	 * Handles internal generation of the solution file
	 * 
	 * @param sheet     Excel sheet for solution
	 * @param timeLimit for finding the solution of the MIP
	 * @param gap       relative primal-to-dual gap as termination criterion
	 * @param address   of user input
	 * @param moreInfo  binary information whether or not to compute skill gaps and
	 *                  diversities
	 * @param pos		position of output row  
	 * @return Excel sheet for solution
	 */
	public static void execute(WritableSheet sheet, int timeLimit, double gap, String address,
			boolean moreInfo, int pos) {
		// Parameters
		double[][] Q_hat;
		double[][] M;
		double[][] Size;
		double[][] E = null;
		int J = -1;
		if(address.startsWith("NumTest")) {
			NumericalTests numTest = new NumericalTests(Integer.parseInt(address.split("_")[1]), Integer.parseInt(address.split("_")[2]), Integer.parseInt(address.split("_")[3]), Integer.parseInt(address.split("_")[4]), Integer.parseInt(address.split("_")[5]));
			Q_hat = numTest.getQ();
			M = numTest.getM();
			Size = numTest.getSize(M);
			if (moreInfo) {
				E = numTest.getE();
				J = (int) Size[4][0];
			}
		}else {
			Size = Converter.arrayListToArray(Converter.excelToArrayList(address, 0, 1, 1));
			Q_hat = Converter.arrayListToArray(Converter.excelToArrayList(address, 1, 1, 1));
			M = Converter.arrayListToArray(Converter.excelToArrayList(address, 3, 1, 1));
			if (moreInfo) {
				E = Converter.arrayListToArray(Converter.excelToArrayList(address, 6, 1, 1));
				J = (int) Size[4][0];
			}
		}
		double[][] Q_bin = Converter.binaryPreferences(Q_hat);
		double[][] Q = Converter.normalizePreferences(Q_hat);
		
		int S = (int) Size[0][0];
		int G = (int) Size[1][0];
		int T = (int) Size[2][0];
	
		double gmin = 0;
		double gmax = S;
		for (int i = 0; i < T; i++) {
			if (M[0][i] > gmin)
				gmin = M[0][i];
			if (M[1][i] < gmax)
				gmax = M[1][i];
		}
		try {
			// Model
			GRBEnv env = new GRBEnv();
			GRBModel model = new GRBModel(env);
			model.set(GRB.StringAttr.ModelName, "Standard Naive");

			// Variables
			GRBVar[][] x = new GRBVar[S][S];

			for (int i = 0; i < S; i++) {
				for (int j = 0; j < S; j++) {
					if (i != j)
						x[i][j] = model.addVar(0, 1, (1/(double)S) * Q_bin[i][j], GRB.BINARY, "x(" + i + ", " + j + ")");
				}
			}

			// Objective function	
			model.set(GRB.IntAttr.ModelSense, GRB.MAXIMIZE);

			// Constraints
			GRBLinExpr expr = new GRBLinExpr();
			for (int i = 0; i < S; i++) {
				expr = new GRBLinExpr();
				for (int j = 0; j < S; j++) {
					if (i != j)
						expr.addTerm(1, x[i][j]);
				}
				model.addConstr(expr, GRB.GREATER_EQUAL, gmin - 1, "Lower bound for group size");
				model.addConstr(expr, GRB.LESS_EQUAL, gmax - 1, "Upper bound for group size");
			}

			for (int i = 0; i < S; i++) {
				for (int j = 0; j < S; j++) {
					for (int k = 0; k < S; k++) {
						if (i != j && j != k && k != i) {
							expr = new GRBLinExpr();
							expr.addTerm(1, x[i][k]);
							expr.addTerm(-1, x[i][j]);
							expr.addTerm(-1, x[j][k]);
							model.addConstr(expr, GRB.GREATER_EQUAL, -1, "Transitivity");
						}
					}
				}
			}

			for (int i = 0; i < S; i++) {
				for (int j = 0; j < S; j++) {
					if (i != j) {
						expr = new GRBLinExpr();
						expr.addTerm(1, x[i][j]);
						expr.addTerm(-1, x[j][i]);
						model.addConstr(expr, GRB.EQUAL, 0, "Symmetry");
					}
				}
			}

			// Solve
			model.set(GRB.DoubleParam.TimeLimit, timeLimit);
			model.set(GRB.DoubleParam.MIPGap, gap);
			model.optimize();

			//Output
			double xsg[][] = Converter.convertFromXtoXsg(x, G);
			double prefExprVal =  model.get(GRB.DoubleAttr.ObjVal);
			System.out.println("Obj: " + prefExprVal);
			
			boolean[] z = new boolean[G];
			for (int j = 0; j < G; j++) {
				for (int k = 0; k < S; k++) {
					if (xsg[k][j]> 1/2)
						z[j] = true;
				}
			}

			String resString = Metrics.getResString(xsg, z);

			double avgSocialHappiness = Metrics.getSocialHappiness(Q, x);
			System.out.println("\nAvg. social satisfaction = " + avgSocialHappiness);

			double[] diversityVal = null;
			double[][] skillGaps = null;
			double skillExprVal = Double.MIN_VALUE;
			
			if (moreInfo) {
				diversityVal = Metrics.getDiverVals(E, x);
				skillGaps = Metrics.getSkillGaps(E, xsg, z);
				skillExprVal = Metrics.getSkillVal(skillGaps, E, x);
			}
			
			// Write
			NumberFormat decimalNo = new NumberFormat("0.00");
			WritableCellFormat numberFormat = new WritableCellFormat(decimalNo);
			if (moreInfo) {
				Number num = new Number(0, pos, prefExprVal, numberFormat);
				sheet.addCell(num);
				num = new Number(1, pos, skillExprVal, numberFormat);
				sheet.addCell(num);
				num = new Number(2, pos, avgSocialHappiness, numberFormat);
				sheet.addCell(num);
				Label label = new Label(4, pos, resString);
				sheet.addCell(label);
				for (int i = 0; i < J; i++) {
					num = new Number(5 + i, pos, skillGaps[1][i], numberFormat);
					sheet.addCell(num);
				}
				for (int i = 0; i < J; i++) {
					num = new Number(5 + J + i, pos, diversityVal[i], numberFormat);
					sheet.addCell(num);
				}
			} else {
				Number num = new Number(0, pos, prefExprVal, numberFormat);
				sheet.addCell(num);
				num = new Number(1, pos, avgSocialHappiness, numberFormat);
				sheet.addCell(num);
				Label label = new Label(3, pos, resString);
				sheet.addCell(label);
			}
			model.dispose();

		} catch (GRBException e) {
			System.out.println("Error code: " + e.getErrorCode() + ". " + e.getMessage());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}