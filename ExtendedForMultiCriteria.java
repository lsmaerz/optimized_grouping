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

public class ExtendedForMultiCriteria {

	public static void main(String[] args) {
	}

	/**
	 * Handles internal generation of the solution file
	 * 
	 * @param sheet     Excel sheet for solutions
	 * @param prio      priority of the skill distribution objective, preference
	 *                  term has prio 0, hence -1,0,1 cover all hierarchical/blended
	 *                  cases
	 * @param tol       tolerance in hierarchical case
	 * @param weight    weight in blended case
	 * @param pos       current row of excel sheet to we written in
	 * @param timeLimit for finding the solution of the MIP
	 * @param gap       relative primal-to-dual gap as termination criterion
	 * @param address   of user input
	 * @return Excel sheet for solutions
	 */
	public static void execute(WritableSheet sheet, int prio, double tol, double weight, int pos, int timeLimit, double gap, String address) {
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
			H = numTest.getH();
			E = numTest.getE();
		} else {
			Size = Converter.arrayListToArray(Converter.excelToArrayList(address, 0, 1, 1));
			Q_hat = Converter.arrayListToArray(Converter.excelToArrayList(address, 1, 1, 1));
			P_hat = Converter.arrayListToArray(Converter.excelToArrayList(address, 2, 1, 1));
			M = Converter.arrayListToArray(Converter.excelToArrayList(address, 3, 1, 1));
			L = Converter.arrayListToArray(Converter.excelToArrayList(address, 4, 1, 1));
			H = Converter.arrayListToArray(Converter.excelToArrayList(address, 5, 1, 1));
			E = Converter.arrayListToArray(Converter.excelToArrayList(address, 6, 1, 1));
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
			model.set(GRB.StringAttr.ModelName, "ExtendedForMultiCriteria");

			// Variables
			GRBVar[][][] a = new GRBVar[S][S][G];
			GRBVar[][][] b = new GRBVar[S][G][T];
			GRBVar[] c_low = model.addVars(J, GRB.CONTINUOUS);
			GRBVar[] c_up = model.addVars(J, GRB.CONTINUOUS);
			GRBVar[][] x = new GRBVar[S][G];
			GRBVar[][] y = new GRBVar[G][T];
			GRBVar[] z = model.addVars(G, GRB.BINARY);

			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					for (int k = 0; k < G; k++) {
						a[i][j][k] = model.addVar(0, 1, 0, GRB.BINARY,
								"a(" + i + ", " + j + ", " + k + ")");
					}
				}
			}

			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					for (int k = 0; k < T; k++)
						b[i][j][k] = model.addVar(0, 1, 0, GRB.BINARY,
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

			// Preference Satisfaction
			GRBLinExpr prefExpr = new GRBLinExpr();
			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					for (int k = 0; k < G; k++) {
						prefExpr.addTerm((L[i][0] * Q[i][j] + L[j][0] * Q[j][i]) * (1/(double)S), a[i][j][k]);
					}
				}
			}
			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					for (int k = 0; k < T; k++)
						prefExpr.addTerm(L[i][1] * P[i][k] * (1/(double)S), b[i][j][k]);
				}
			}

			// Skill Distribution
			GRBLinExpr skillExpr = new GRBLinExpr();
			double tmpMu;
			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					for (int k = 0; k < G; k++) {
						tmpMu = 0;
						for (int l = 0; l < J; l++) {
							tmpMu += E[1][l] * Math.abs(E[2 + i][l] - E[2 + j][l]);
						}
						skillExpr.addTerm(tmpMu, a[i][j][k]);
					}
				}
			}
			for (int i = 0; i < J; i++) {
				skillExpr.addTerm(-E[0][i], c_low[i]);
				skillExpr.addTerm(E[0][i], c_up[i]);
			}
			
			model.set(GRB.IntAttr.ModelSense, GRB.MAXIMIZE);
			if (prio == 0) {
				model.setObjectiveN(prefExpr, 0, 0, 1, 0, 0, "Preference Satisfaction");
				model.setObjectiveN(skillExpr, 1, 0, weight, 0, 0, "Skill distribution");
				//model.set(GRB.IntParam.GomoryPasses, 0);
				//model.set(GRB.DoubleParam.Heuristics, 0.001);
				//model.set(GRB.IntParam.PreDual, 1);
			}else if (prio < 0) {
				GRBEnv e0 = model.getMultiobjEnv(0);
				e0.set(GRB.DoubleParam.TimeLimit, ((double) timeLimit)/5);
				GRBEnv e1 = model.getMultiobjEnv(1);
				e1.set(GRB.DoubleParam.TimeLimit, ((double) 2*timeLimit)/5);
				model.setObjectiveN(prefExpr, 0, 1, 1, 0, tol, "Preference Satisfaction I/II");
				model.setObjectiveN(skillExpr, 1, 0, 1, 0, 0, "Skill distribution");
				model.setObjectiveN(prefExpr, 2, -1, 1, 0, 0, "Preference Satisfaction II/II");
				//model.set(GRB.IntParam.GomoryPasses, 0);
				//model.set(GRB.IntParam.VarBranch, 1);
			}else if (prio > 0) {
				GRBEnv e0 = model.getMultiobjEnv(0);
				e0.set(GRB.DoubleParam.TimeLimit, ((double) timeLimit)/5);
				GRBEnv e1 = model.getMultiobjEnv(1);
				e1.set(GRB.DoubleParam.TimeLimit, ((double) 2*timeLimit)/5);
				model.setObjectiveN(skillExpr, 0, 1, 1, 0, tol, "Skill distribution I/II");
				model.setObjectiveN(prefExpr, 1, 0, 1, 0, 0, "Preference Satisfaction");
				model.setObjectiveN(skillExpr, 2, -1, 1, 0, 0, "Skill distribution II/II");
				//model.set(GRB.IntParam.CutPasses, 3);
				//model.set(GRB.IntParam.Aggregate, 2);
				//model.set(GRB.DoubleParam.Heuristics, 0);
			}

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

			for (int i = 0; i < G; i++) {
				for (int j = 0; j < I; j++) {
					GRBLinExpr expr = new GRBLinExpr();
					for (int k = 0; k < S; k++) {
						expr.addTerm(H[2 + k][j], x[k][i]);
					}
					model.addConstr(expr, GRB.LESS_EQUAL, H[1][j], "Hard skills upper bound");
					expr.addTerm(-S, z[i]);
					model.addConstr(expr, GRB.GREATER_EQUAL, H[0][j] - S, "Hard skills lower bound");
				}
			}

			for (int i = 0; i < G; i++) {
				for (int j = 0; j < J; j++) {
					GRBLinExpr expr = new GRBLinExpr();
					for (int k = 0; k < S; k++) {
						expr.addTerm(E[2 + k][j], x[k][i]);
					}
					expr.addTerm(-1, c_low[j]);
					expr.addTerm(-S, z[i]);
					model.addConstr(expr, GRB.GREATER_EQUAL, -S, "Experiences lower bound");
				}
			}

			for (int i = 0; i < G; i++) {
				for (int j = 0; j < J; j++) {
					GRBLinExpr expr = new GRBLinExpr();
					for (int k = 0; k < S; k++) {
						expr.addTerm(E[2 + k][j], x[k][i]);
					}
					expr.addTerm(-1, c_up[j]);
					model.addConstr(expr, GRB.LESS_EQUAL, 0, "Experiences upper bound");
				}
			}

			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					GRBLinExpr expr = new GRBLinExpr();
					expr.addTerm(1, x[i][j]);
					expr.addTerm(-1, z[j]);
					model.addConstr(expr, GRB.LESS_EQUAL, 0, "Group used?");
				}
			}

			// Solve
			//model.tune();
			model.set(GRB.DoubleParam.TimeLimit, timeLimit);
			model.set(GRB.DoubleParam.MIPGap, gap);
			model.set(GRB.DoubleParam.ImproveStartTime, ((double) 65*timeLimit)/100);
			model.optimize();

			// Output
			String resString = Metrics.getResString(x, y, z);

			double avgSocialHappiness = Metrics.getSocialHappiness(Q, a);
			System.out.println("\nAvg. social satisfaction = " + avgSocialHappiness);

			double avgTopicHappiness = Metrics.getTopicHappiness(P, b);
			System.out.println("\nAvg. topic satisfaction = " + avgTopicHappiness);

			double[][] skillGaps = Metrics.getSkillGaps(E, x, z);

			//double prefExprVal = Metrics.getPrefVal(G, Q, P, L, a, b);
			double prefExprVal = prefExpr.getValue();
			System.out.println("\nPreference Satisfacion Value: " + prefExprVal);
			
			//double skillExprVal = Metrics.getSkillVal(skillGaps, E, a);
			double skillExprVal = skillExpr.getValue();
			System.out.println("\nSkill Distribution Value: " + skillExprVal);

			double[] diversityVal = Metrics.getDiverVals(E, a);

			NumberFormat decimalNo = new NumberFormat("0.00");
			WritableCellFormat numberFormat = new WritableCellFormat(decimalNo);			
			
			Number num = new Number(2, pos, prefExprVal, numberFormat);
			sheet.addCell(num);
			num = new Number(3, pos, skillExprVal, numberFormat);
			sheet.addCell(num);
			num = new Number(4, pos, avgSocialHappiness, numberFormat);
			sheet.addCell(num);
			num = new Number(5, pos, avgTopicHappiness, numberFormat);
			sheet.addCell(num);
			Label label = new Label(7, pos, resString);
			sheet.addCell(label);
			for (int i = 0; i < J; i++) {
				num = new Number(8 + i, pos, skillGaps[1][i], numberFormat);
				sheet.addCell(num);
			}
			for (int i = 0; i < J; i++) {
				num = new Number(8 + J + i, pos, diversityVal[i], numberFormat);
				sheet.addCell(num);
			}
			model.dispose();
		} catch (GRBException e) {
			System.out.println("Error code: " + e.getErrorCode() + ". " + e.getMessage());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}