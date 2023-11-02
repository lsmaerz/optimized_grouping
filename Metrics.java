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

public class Metrics {
	
	/**
	 * Computes average social happiness
	 * @param Q matrix Q from user input
	 * @param a matrix a_{ssg} from the paper
	 * @return average social happiness
	 */
	public static double getSocialHappiness (double[][] Q, GRBVar[][][] a) {
		try {
			int S = a.length;
			int G = a[0][0].length;
			double res = 0;
			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					for (int k = 0; k < G; k++) {
						res += (1/(double)S) * (Q[i][j] + Q[j][i]) * a[i][j][k].get(GRB.DoubleAttr.X);
					}
				}
			}
			return res;
		} catch (Exception e) {
			e.printStackTrace();
			return Double.MIN_VALUE;
		}
	}
	
	/**
	 * Computes average social happiness
	 * @param Q matrix Q from user input
	 * @param x matrix x_{ss} from the paper
	 * @return average social happiness
	 */
	public static double getSocialHappiness (double[][] Q, GRBVar[][] x) {
		try {
			int S = x.length;
			double sum = 0;
			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					if (x[i][j].get(GRB.DoubleAttr.X) > 1 / 2) {
						sum += Q[i][j] + Q[j][i];
					}
				}
			}
			return sum / S;
		} catch (Exception e) {
			e.printStackTrace();
			return Double.MIN_VALUE;
		}
	}
	
	/**
	 * Computes average topic happiness
	 * @param P matrix P from user input
	 * @param b matrix b_{sgt} from the paper
	 * @return average topic happiness
	 */
	public static double getTopicHappiness (double[][] P, GRBVar[][][] b) {
		try {
			int S = b.length;
			int G = b[0].length;
			int T = b[0][0].length;
			double res = 0;
			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					for (int k = 0; k < T; k++) {
						res += (1/(double)S) * P[i][k] * b[i][j][k].get(GRB.DoubleAttr.X);
					}
				}
			}
			return res;
		} catch (Exception e) {
			e.printStackTrace();
			return Double.MIN_VALUE;
		}
	}
	
	/**
	 * Computes skill gaps between collective group skills
	 * @param E matrix E from user input
	 * @param x matrix x_{ss} from the paper
	 * @param z vector z_g from the paper
	 * @return skill gaps
	 */
	public static double[][] getSkillGaps (double E[][], GRBVar[][] x, GRBVar[] z) {
		try {
			int S = x.length;
			int G = x[0].length;
			int J = E[0].length;
			double[][] skillGaps = new double[2][J];
			for (int i = 0; i < J; i++) {
				double low = S;
				double up = 0;
				for (int j = 0; j < G; j++) {
					if (z[j].get(GRB.DoubleAttr.X) > 1 / 2) {
						double tmp = 0;
						for (int k = 0; k < S; k++) {
							tmp += x[k][j].get(GRB.DoubleAttr.X) * E[2 + k][i];
						}
						if (tmp < low)
							low = tmp;
						if (tmp > up)
							up = tmp;
					}
				}
				skillGaps[0][i] = up - low;
				skillGaps[1][i] = low / up;
			}
			return skillGaps;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Computes skill gaps between collective group skills
	 * @param E matrix E from user input
	 * @param x matrix x_{ss} from the paper
	 * @param z vector z_g from the paper
	 * @return skill gaps
	 */
	public static double[][] getSkillGaps (double E[][], GRBVar[][] x, boolean[] z) {
		try {
			int S = x.length;
			int G = x[0].length;
			int J = E[0].length;
			double[][] skillGaps = new double[2][J];
			for (int i = 0; i < J; i++) {
				double low = S;
				double up = 0;
				for (int j = 0; j < G; j++) {
					if (z[j]) {
						double tmp = 0;
						for (int k = 0; k < S; k++) {
							tmp += x[k][j].get(GRB.DoubleAttr.X) * E[2 + k][i];
						}
						if (tmp < low)
							low = tmp;
						if (tmp > up)
							up = tmp;
					}
				}
				skillGaps[0][i] = up - low;
				skillGaps[1][i] = low / up;
			}
			return skillGaps;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Computes skill gaps between collective group skills
	 * @param E matrix E from user input
	 * @param x matrix x_{ss} from the paper
	 * @param z vector z_g from the paper
	 * @return skill gaps
	 */
	public static double[][] getSkillGaps (double E[][], double[][] x, boolean[] z) {
		try {
			int S = x.length;
			int G = x[0].length;
			int J = E[0].length;
			double[][] skillGaps = new double[2][J];
			for (int i = 0; i < J; i++) {
				double low = S;
				double up = 0;
				for (int j = 0; j < G; j++) {
					if (z[j]) {
						double tmp = 0;
						for (int k = 0; k < S; k++) {
							tmp += x[k][j] * E[2 + k][i];
						}
						if (tmp < low)
							low = tmp;
						if (tmp > up)
							up = tmp;
					}
				}
				skillGaps[0][i] = up - low;
				skillGaps[1][i] = low / up;
			}
			return skillGaps;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Computes preference value
	 * @param G value G from paper
	 * @param Q	matrix Q from user input
	 * @param P matrix P from user input
	 * @param L matrix L from user input
	 * @param a matrix a_{ssg} from paper
	 * @param b matrix b_{sgt} from papaer
	 * @return preference value
	 */
	
	public static double getPrefVal (int G, double[][] Q, double[][] P, double[][] L, GRBVar[][][] a, GRBVar[][][] b){
		try {
			int S = Q.length;
			int T = P[0].length;
			double res = 0;
			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					for (int k = 0; k < G; k++) {
						res += (L[i][0] * Q[i][j] + L[j][0] * Q[j][i]) * (1/(double)S) * a[i][j][k].get(GRB.DoubleAttr.X);
					}
				}
			}
			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					for (int k = 0; k < T; k++)
						res += L[i][1] * P[i][k] * (1/(double)S) * b[i][j][k].get(GRB.DoubleAttr.X);
				}
			}
			return res;
		}catch(Exception e) {
			e.printStackTrace();
			return Double.MIN_VALUE;
		}
	}
	
	
	/**
	 * Computes preference value
	 * @param G value G from paper
	 * @param Q	matrix Q from user input
	 * @param P matrix P from user input
	 * @param L matrix L from user input
	 * @param x matrix x_{ss} from paper
	 * @param b matrix b_{sgt} from paper
	 * @return preference value
	 */
	public static double getPrefVal (int G, double[][] Q, double[][] P, double[][] L, double[][] x, GRBVar[][][] b){
		try {
			double prefExpr = 0;
			int S = Q.length;
			int T = P[0].length;
			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					prefExpr += (L[i][0] * Q[i][j] + L[j][0] * Q[j][i]) * (1/(double)S) * x[i][j];
				}
			}
			for (int i = 0; i < S; i++) {
				for (int j = 0; j < G; j++) {
					for (int k = 0; k < T; k++)
						prefExpr += L[i][1] * P[i][k] * (1/(double)S) * b[i][j][k].get(GRB.DoubleAttr.X);
				}
			}
			return prefExpr;
		}catch(Exception e) {
			e.printStackTrace();
			return Double.MIN_VALUE;
		}
	}
	
	/**
	 * Computes skill values
	 * @param skillGaps between collective group skills
	 * @param E matrix E from user input
	 * @param a matrix a_{ssg} from the paper
	 * @return skill values
	 */
	
	public static double getSkillVal (double[][] skillGaps, double E[][], GRBVar[][][] a) {
		try {
			int S = a.length;
			int G = a[0][0].length;
			int J = E[0].length;
			double res = 0;
			double tmpMu;
			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					for (int k = 0; k < G; k++) {
						tmpMu = 0;
						for (int l = 0; l < J; l++) {
							tmpMu += E[1][l] * Math.abs(E[2 + i][l] - E[2 + j][l]);
						}
						res += tmpMu * a[i][j][k].get(GRB.DoubleAttr.X);
					}
				}
			}
			for (int i = 0; i < J; i++) {
				res += E[0][i] * skillGaps[0][i];
			}
			return res;
		} catch (Exception e) {
			e.printStackTrace();
			return Double.MIN_VALUE;
		}
	}
	
	
	/**
	 * Computes skill values
	 * @param skillGaps between collective group skills
	 * @param E matrix E from user input
	 * @param x	matrix x_{ss} from the paper
	 * @return skill values
	 */
	public static double getSkillVal (double[][] skillGaps, double E[][], GRBVar[][] x) {
		try {
			int S = x.length;
			int J = E[0].length;
			double skillExprVal = 0;
			for (int i = 0; i < S; i++) {
				for (int j = i + 1; j < S; j++) {
					double tmpMu = 0;
					for (int l = 0; l < J; l++) {
						tmpMu += E[1][l] * Math.abs(E[2 + i][l] - E[2 + j][l]);
					}
					skillExprVal += tmpMu * x[i][j].get(GRB.DoubleAttr.X);
				}
			}
			for (int i = 0; i < J; i++) {
				skillExprVal += E[0][i] * skillGaps[0][i];
			}
			return skillExprVal;
		} catch (Exception e) {
			e.printStackTrace();
			return Double.MIN_VALUE;
		}
	}

	/**
	 * Computes diversity values
	 * @param E matrix E from user input
	 * @param a matrix a_{ssg} from the paper
	 * @return diversity values
	 */
	public static double[] getDiverVals (double E[][], GRBVar[][][] a) {
		try {
			int S = a.length;
			int G = a[0][0].length;
			int J = E[0].length;
			double count = 0;
			double[] diversityVal = new double[J];
			for (int j = 0; j < S; j++) {
				for (int k = j + 1; k < S; k++) {
					for (int l = 0; l < G; l++) {
						count += a[j][k][l].get(GRB.DoubleAttr.X);
						for (int i = 0; i < J; i++) {
							diversityVal[i] += a[j][k][l].get(GRB.DoubleAttr.X) *
									Math.abs(E[2 + j][i] - E[2 + k][i]);
						}
					}
				}
			}
			for (int i = 0; i < J; i++) {
				diversityVal[i] = diversityVal[i] / count;
			}
			return diversityVal;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Computes diversity values
	 * @param E matrix E from user input
	 * @param x	matrix x_{ss} from the paper
	 * @return diversity values
	 */
	public static double[] getDiverVals (double E[][], GRBVar[][] x) {
		try {
			int S = x.length;
			int J = E[0].length;
			double[] diversityVal = new double[J];
			double count = 0;				
			for (int j = 0; j < S; j++) {
				for (int k = j + 1; k < S; k++) {
					count += x[j][k].get(GRB.DoubleAttr.X);
					for (int i = 0; i < J; i++) {
						diversityVal[i] += x[j][k].get(GRB.DoubleAttr.X) * Math.abs(E[2 + j][i] - E[2 + k][i]);
					}
				}
			}
			for (int i = 0; i < J; i++) {
				diversityVal[i] = diversityVal[i] / count;
			}
			return diversityVal;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Queries student-group-topic assignment as string
	 * @param x matrix x_{sg} from the paper
	 * @param y matrix y_{gt} from the paper
	 * @param z vector z_g from the paper
	 * @return group assignment as string
	 */
	public static String getResString (GRBVar[][] x, GRBVar[][] y, GRBVar[] z) {
		try {
			int S = x.length;
			int G = x[0].length;
			int T = y[0].length;
			String resString = "";
			for (int g = 0; g < G; g++) {
				if (z[g].get(GRB.DoubleAttr.X) > 1 / 2) {
					System.out.println();
					System.out.println("Group " + g);
					resString = resString + "#G# ";
					for (int t = 0; t < T; t++) {
						if (y[g][t].get(GRB.DoubleAttr.X) == 1) {
							System.out.println("Topic " + t);
							resString = resString + "T" + t + ": ";
						}

					}
					for (int s = 0; s < S; s++) {
						if (x[s][g].get(GRB.DoubleAttr.X) == 1) {
							System.out.println("Student " + s);
							resString = resString + "S" + s + ". ";
						}

					}
				}
			}
			return resString;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Queries student-group-topic assignment as string
	 * @param x matrix x_{sg} from the paper
	 * @param y matrix y_{gt} from the paper
	 * @param z vector z_g from the paper
	 * @return group assignment as string
	 */
	public static String getResString (GRBVar[][] x, GRBVar[][] y, boolean[] z) {
		try {
			int S = x.length;
			int G = x[0].length;
			int T = y[0].length;
			String resString = "";
			for (int g = 0; g < G; g++) {
				if (z[g]) {
					System.out.println();
					System.out.println("Group " + g);
					resString = resString + "#G# ";
					for (int t = 0; t < T; t++) {
						if (y[g][t].get(GRB.DoubleAttr.X) == 1) {
							System.out.println("Topic " + t);
							resString = resString + "T" + t + ": ";
						}

					}
					for (int s = 0; s < S; s++) {
						if (x[s][g].get(GRB.DoubleAttr.X) == 1) {
							System.out.println("Student " + s);
							resString = resString + "S" + s + ". ";
						}

					}
				}
			}
			return resString;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Queries student-group assignment as string
	 * @param x matrix x_{sg} from the paper
	 * @param y matrix y_{gt} from the paper
	 * @param z vector z_g from the paper
	 * @return group assignment as string
	 */
	public static String getResString (double[][] x, boolean[] z) {
		try {
			int S = x.length;
			int G = x[0].length;
			String resString = "";
			for (int g = 0; g < G; g++) {
				if (z[g]) {
					String tmp = "";
					for (int s = 0; s < S; s++) {
						if (x[s][g] > 1/2) {
							tmp = tmp + "S" + s + ". ";
						}
					}
					System.out.println("\nGroup " + g);
					System.out.println(tmp);
					resString = resString + "#G# " + tmp;
				}
			}
			return resString;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
}
