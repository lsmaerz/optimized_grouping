import java.util.ArrayList;
import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;
import gurobi.*;

public class ExtendedForTopicAssignment {

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
	 * @param pos       position of output row
	 * @return Excel sheet for solution
	 */
	public static void execute(WritableSheet sheet, int timeLimit, double gap, String address,
			boolean moreInfo, int pos) {
		// Parameters
		double[][] Q_hat;
		double[][] P_hat;
		double[][] M;
		double[][] L;
		double[][] Size;
		double[][] H = null;
		double[][] E = null;
		if (address.startsWith("NumTest")) {
			NumericalTests numTest = new NumericalTests(Integer.parseInt(address.split("_")[1]),
					Integer.parseInt(address.split("_")[2]), Integer.parseInt(address.split("_")[3]),
					Integer.parseInt(address.split("_")[4]), Integer.parseInt(address.split("_")[5]));
			Q_hat = numTest.getQ();
			P_hat = numTest.getP();
			M = numTest.getM();
			L = numTest.getL();
			Size = numTest.getSize(M);
			if (moreInfo) {
				H = numTest.getH();
				E = numTest.getE();
			}
		} else {
			Size = Converter.arrayListToArray(Converter.excelToArrayList(address, 0, 1, 1));
			Q_hat = Converter.arrayListToArray(Converter.excelToArrayList(address, 1, 1, 1));
			P_hat = Converter.arrayListToArray(Converter.excelToArrayList(address, 2, 1, 1));
			M = Converter.arrayListToArray(Converter.excelToArrayList(address, 3, 1, 1));
			L = Converter.arrayListToArray(Converter.excelToArrayList(address, 4, 1, 1));
			if (moreInfo) {
				H = Converter.arrayListToArray(Converter.excelToArrayList(address, 5, 1, 1));
				E = Converter.arrayListToArray(Converter.excelToArrayList(address, 6, 1, 1));
			}
		}
		double[][] Q = Converter.normalizePreferences(Q_hat);
		double[][] P = Converter.normalizePreferences(P_hat);

		int S = (int) Size[0][0];
		int G = (int) Size[1][0];
		int T = (int) Size[2][0];
		int I = (int) Size[3][0];
		int J = (int) Size[4][0];

		try {
			// Model
			GRBEnv env = new GRBEnv();
			GRBModel model = new GRBModel(env);
			model.set(GRB.StringAttr.ModelName, "ExtendedForTopicAssignment");

			// Variables
			GRBVar[][][] a = new GRBVar[S][S][G];
			GRBVar[][][] b = new GRBVar[S][G][T];
			GRBVar[][] x = new GRBVar[S][G];
			GRBVar[][] y = new GRBVar[G][T];

			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					for (int k = 0; k < G; k++)
						a[i][j][k] = model.addVar(0, 1, (1/(double)S)*(L[i][0] * Q[i][j] + L[j][0] * Q[j][i]), GRB.BINARY,
								"a(" + i + ", " + j + ", " + k + ")");
				}
			}

			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					for (int k = 0; k < T; k++)
						b[i][j][k] = model.addVar(0, 1, L[i][1] * P[i][k] * (1/(double)S), GRB.BINARY,
								"b(" + i + ", " + j + ", " + k + ")");
				}
			}

			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					x[i][j] = model.addVar(0, 1, 0, GRB.BINARY, "x(" + i + ", " + j + ")");
				}
			}
			for (int i = 0; i < G; i++) {
				for (int j = 0; j < T; j++) {
					y[i][j] = model.addVar(0, 1, 0, GRB.BINARY, "y(" + i + ", " + j + ")");
				}
			}

			// Objective
			model.set(GRB.IntAttr.ModelSense, GRB.MAXIMIZE);

			// Constraints
			for (int i = 0; i < S; i++) {
				GRBLinExpr expr = new GRBLinExpr();
				for (int j = 0; j < G; j++) {
					expr.addTerm(1, x[i][j]);
				}
				model.addConstr(expr, GRB.EQUAL, 1, "Everyone is in one group");
			}

			for (int i = 0; i < G; i++) {
				GRBLinExpr expr = new GRBLinExpr();
				for (int j = 0; j < T; j++) {
					expr.addTerm(1, y[i][j]);
				}
				model.addConstr(expr, GRB.LESS_EQUAL, 1, "Groups have at most one topic");
			}

			for (int i = 0; i < G; i++) {
				GRBLinExpr expr = new GRBLinExpr();
				for (int j = 0; j < T; j++) {
					expr.addTerm(1, y[i][j]);
				}
				for (int j = 0; j < S; j++) {
					expr.addTerm(-1, x[j][i]);
				}
				model.addConstr(expr, GRB.LESS_EQUAL, 0, "Only nonempty groups may have topic");
			}

			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					GRBLinExpr expr = new GRBLinExpr();
					for (int k = 0; k < T; k++) {
						expr.addTerm(1, y[j][k]);
					}
					expr.addTerm(-1, x[i][j]);
					model.addConstr(expr, GRB.GREATER_EQUAL, 0, "Nonempty groups must have a topic");
				}
			}

			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					for (int k = 0; k < G; k++) {
						GRBLinExpr expr = new GRBLinExpr();
						expr.addTerm(1, a[i][j][k]);
						expr.addTerm(-0.5, x[i][k]);
						expr.addTerm(-0.5, x[j][k]);
						model.addConstr(expr, GRB.LESS_EQUAL, 0, "Be together in a certain group I/II");
					}

				}
			}

			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					for (int k = 0; k < G; k++) {
						GRBLinExpr expr = new GRBLinExpr();
						expr.addTerm(1, a[i][j][k]);
						expr.addTerm(-1, x[i][k]);
						expr.addTerm(-1, x[j][k]);
						model.addConstr(expr, GRB.GREATER_EQUAL, -1.5, "Be together in a certain group II/II");
					}

				}
			}

			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					for (int k = 0; k < T; k++) {
						GRBLinExpr expr = new GRBLinExpr();
						expr.addTerm(1, b[i][j][k]);
						expr.addTerm(-0.5, x[i][j]);
						expr.addTerm(-0.5, y[j][k]);
						model.addConstr(expr, GRB.LESS_EQUAL, 0, "Have certain topic via certain group I/II");
					}

				}
			}

			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					for (int k = 0; k < T; k++) {
						GRBLinExpr expr = new GRBLinExpr();
						expr.addTerm(1, b[i][j][k]);
						expr.addTerm(-1, x[i][j]);
						expr.addTerm(-1, y[j][k]);
						model.addConstr(expr, GRB.GREATER_EQUAL, -1.5, "Have certain topic via certain group II/II");
					}

				}
			}

			for (int i = 0; i < G; i++) {
				for (int j = 0; j < T; j++) {
					GRBLinExpr expr = new GRBLinExpr();
					for (int k = 0; k < S; k++) {
						expr.addTerm(1, b[k][i][j]);
					}
					expr.addTerm(-M[0][j], y[i][j]);
					model.addConstr(expr, GRB.GREATER_EQUAL, 0, "Minimum group size for certain topic");
				}
			}

			for (int i = 0; i < G; i++) {
				for (int j = 0; j < T; j++) {
					GRBLinExpr expr = new GRBLinExpr();
					for (int k = 0; k < S; k++) {
						expr.addTerm(1, b[k][i][j]);
					}
					expr.addTerm(-M[1][j], y[i][j]);
					model.addConstr(expr, GRB.LESS_EQUAL, 0, "Maximum group size for certain topic");
				}
			}

			for (int i = 0; i < T; i++) {
				GRBLinExpr expr = new GRBLinExpr();
				for (int j = 0; j < G; j++) {
					expr.addTerm(1, y[j][i]);
				}
				model.addConstr(expr, GRB.GREATER_EQUAL, M[2][i], "Minimum occurrences of certain topic");
			}

			for (int i = 0; i < T; i++) {
				GRBLinExpr expr = new GRBLinExpr();
				for (int j = 0; j < G; j++) {
					expr.addTerm(1, y[j][i]);
				}
				model.addConstr(expr, GRB.LESS_EQUAL, M[3][i], "Maximum occurrences of certain topic");
			}

			// Solve
			model.set(GRB.DoubleParam.TimeLimit, timeLimit);
			model.set(GRB.DoubleParam.MIPGap, gap);
			model.optimize();

			// Output
			boolean[] z = new boolean[G];
			for (int j = 0; j < G; j++) {
				for (int k = 0; k < S; k++) {
					if (x[k][j].get(GRB.DoubleAttr.X) > 1 / 2)
						z[j] = true;
				}
			}
			String resString = Metrics.getResString(x, y, z);
			
			double avgSocialHappiness = Metrics.getSocialHappiness(Q, a);
			System.out.println("\nAvg. social satisfaction = " + avgSocialHappiness);

			double avgTopicHappiness = Metrics.getTopicHappiness(P, b);
			System.out.println("\nAvg. topic satisfaction = " + avgTopicHappiness);

			double prefExprVal = model.get(GRB.DoubleAttr.ObjVal);
			System.out.println("\nPreference Satisfaction Value: " + prefExprVal);

			double[] diversityVal = null;
			double[][] skillGaps = null;
			double skillExprVal = Double.MIN_VALUE;
			
			if (moreInfo) {
				diversityVal = Metrics.getDiverVals(E, a);				
				skillGaps = Metrics.getSkillGaps(E, x, z);
				skillExprVal = Metrics.getSkillVal(skillGaps, E, a);
			}

			NumberFormat decimalNo = new NumberFormat("0.00");
			WritableCellFormat numberFormat = new WritableCellFormat(decimalNo);
			if (moreInfo) {
				Number num = new Number(0, pos, prefExprVal, numberFormat);
				sheet.addCell(num);
				num = new Number(1, pos, skillExprVal, numberFormat);
				sheet.addCell(num);
				num = new Number(2, pos, avgSocialHappiness, numberFormat);
				sheet.addCell(num);
				num = new Number(3, pos, avgTopicHappiness, numberFormat);
				sheet.addCell(num);
				Label label = new Label(5, pos, resString);
				sheet.addCell(label);
				for (int i = 0; i < J; i++) {
					num = new Number(6 + i, pos, skillGaps[1][i], numberFormat);
					sheet.addCell(num);
				}
				for (int i = 0; i < J; i++) {
					num = new Number(6 + J + i, pos, diversityVal[i], numberFormat);
					sheet.addCell(num);
				}
			} else {
				Number num = new Number(0, pos, prefExprVal, numberFormat);
				sheet.addCell(num);
				num = new Number(1, pos, avgSocialHappiness, numberFormat);
				sheet.addCell(num);
				num = new Number(2, pos, avgTopicHappiness, numberFormat);
				sheet.addCell(num);
				Label label = new Label(4, pos, resString);
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