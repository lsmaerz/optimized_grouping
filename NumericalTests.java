import java.io.File;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.util.Random;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class NumericalTests {
	private Random rand;
	private int S, T, I, J;
	final private double hateSocial = 0.03;
	final private int maxPeerSize = 4;
	final private double socialShare = 0.65;
	final private double hateTopic = 0.25;
	final private double likeTopic = 0.25;
	final static int N = 100;
	final static int N_multi = 50;
	final static int N_frontier = 1;
	final static int maxRuntime = 1800*17; //for 17 frontier approximations divided in 17 equal parts, 30 min each
	final static double dualityGap = 0.1;
	
	/**
	 * Constructor of this class.
	 * @param seed	determines all random values.
	 * @param S		number of students
	 * @param T		number of topics
	 * @param I		number of hard skills
	 * @param J		number of experience skills
	 */
	public NumericalTests(long seed, int S, int T, int I, int J) {
		rand = new Random(seed);
		this.S = S;
		this.T = T;
		this.I = I;
		this.J = J;
	}
	
	/**
	 * Calls all numerical tests.
	 * @param args	(not used)
	 */
	public static void main(String[] args) {	
		NumericalTests.changeLogFile("D:\\Studium\\Semester_9\\Masterseminar_OR\\ConsoleOutput.txt");
		
		for (int i = 1; i < 4; i++) {
			NumericalTests.testNaive(10*i, 1, 2, 3, "D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\Results\\results_naive_N" + N + "_S" + 10*i + "_Gap" + dualityGap + ".xls");
			NumericalTests.testImproved(10*i, 1, 2, 3, "D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\Results\\results_improved_N" + N + "_S" + 10*i + "_Gap" + dualityGap + ".xls");
			for (int j = 1; j < 4; j++) {
				//NumericalTests.testSingle(10*i, 5*j, 2, 3, "D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\Results\\results_single_N" + N + "_S" + 10*i + "_T" + 5*j + "_Gap" + dualityGap + ".xls");
				//NumericalTests.testMulti(10*i, 5*j, 2, 3, "D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\Results\\results_multi_N" + N + "_S" + 10*i + "_T" + 5*j + "_Gap" + dualityGap + ".xls");
				//NumericalTests.testFrontier(10*i, 5*j, 2, 3, "D:\\Studium\\Semester_9\\Masterseminar_OR\\results_frontier_N" + N_frontier + "_S" + 10*i + "_T" + 5*j + "_Gap" + dualityGap + ".xls");
			}
		}
	}
	
	/**
	 * Redirects console output to given address
	 * @param address	destination
	 */
	public static void changeLogFile(String address) {
		try {
			PrintStream out = new PrintStream(new FileOutputStream(address));
			System.setOut(out);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Tests StandardNaive numerically.
	 * @param S_hat			S of instance
	 * @param T_hat			T of instance
	 * @param I_hat			I of instance
	 * @param J_hat			J of instance
	 * @param outputAddress	address for results
	 */
	public static void testNaive(int S_hat, int T_hat, int I_hat, int J_hat, String outputAddress) {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(outputAddress));
			WritableSheet sheet = workbook.createSheet("Results", 0);
			System.out.println("---------------Naive-----------------");
			for (int i = 0; i < N; i++) {
				System.out.println("///////Naive i: " + i);
				String Instance = "NumTest_" + i + "_" + S_hat + "_" + T_hat + "_" + I_hat + "_" + J_hat;
				Wrapper.naive(Instance, outputAddress, maxRuntime, dualityGap, true, i+1, sheet);
			}
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Tests StandardImproved numerically.
	 * @param S_hat			S of instance
	 * @param T_hat			T of instance
	 * @param I_hat			I of instance
	 * @param J_hat			J of instance
	 * @param outputAddress	address for results
	 */
	public static void testImproved(int S_hat, int T_hat, int I_hat, int J_hat, String outputAddress) {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(outputAddress));
			WritableSheet sheet = workbook.createSheet("Results", 0);
			System.out.println("---------------Improved-----------------");
			for (int i = 0; i < N; i++) {
				System.out.println("///////Improved i: " + i);
				String Instance = "NumTest_" + i + "_" + S_hat + "_" + T_hat + "_" + I_hat + "_" + J_hat;
				Wrapper.improved(Instance, outputAddress, maxRuntime, dualityGap, true, i+1, sheet);
			}
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Tests ExtendedForTopicAssignment numerically.
	 * @param S_hat			S of instance
	 * @param T_hat			T of instance
	 * @param I_hat			I of instance
	 * @param J_hat			J of instance
	 * @param outputAddress	address for results
	 */
	public static void testSingle(int S_hat, int T_hat, int I_hat, int J_hat, String outputAddress) {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(outputAddress));
			WritableSheet sheet = workbook.createSheet("Results", 0);
			System.out.println("---------------Single-----------------");
			for (int i = 0; i < N; i++) {
				System.out.println("///////Single i: " + i);
				String Instance = "NumTest_" + i + "_" + S_hat + "_" + T_hat + "_" + I_hat + "_" + J_hat;
				Wrapper.single(Instance, outputAddress, maxRuntime, dualityGap, true, i+1, sheet);
			}
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Tests ExtendedForMultiCriteria numerically.
	 * @param S_hat			S of instance
	 * @param T_hat			T of instance
	 * @param I_hat			I of instance
	 * @param J_hat			J of instance
	 * @param outputAddress	address for results
	 */
	public static void testMulti(int S_hat, int T_hat, int I_hat, int J_hat, String outputAddress) {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(outputAddress));
			WritableSheet sheet = workbook.createSheet("Results", 0);
			System.out.println("---------------Multi-----------------");
			for (int i = 0; i < N_multi; i++) {
				System.out.println("//////////Multi: i = " + i);
				String Instance = "NumTest_" + i + "_" + S_hat + "_" + T_hat + "_" + I_hat + "_" + J_hat;
				int j = -1 + i%3;
				Wrapper.multi(Instance, outputAddress, maxRuntime, j, j == 0 ? 0.75 + (i%2) * 0.5 : 0.15 + (i%2) * 0.15, dualityGap, i+1, sheet);
			}
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Tests approximation of efficient frontier numerically.
	 * @param S_hat			S of instance
	 * @param T_hat			T of instance
	 * @param I_hat			I of instance
	 * @param J_hat			J of instance
	 * @param outputAddress	address for results
	 */
	public static void testFrontier(int S_hat, int T_hat, int I_hat, int J_hat, String outputAddress) {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(outputAddress));
			WritableSheet[] sheets = new WritableSheet[N_frontier];
			System.out.println("---------------Frontier-----------------");
			for (int i = 0; i < N_frontier; i++) {
				System.out.println("//////////Frontier: i = " + i);
				String Instance = "NumTest_" + i + "_" + S_hat + "_" + T_hat + "_" + I_hat + "_" + J_hat;
				sheets[i] = workbook.createSheet("Results_" + i, i);
				Wrapper.frontier(Instance, outputAddress, maxRuntime, dualityGap, sheets[i]);
			}
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Prints the randomly generated artificial user inputs.
	 */
	public void printAll() {
		double[][] Q = this.getQ();
		double[][] P = this.getP();
		double[][] M = this.getM();
		double[][] H = this.getH();
		double[][] E = this.getE();
		double[][] Size = this.getSize(M);
		
		System.out.println("\nSize:");
		for (int i = 0; i < 5; i++) {
			System.out.println(Size[i][0]);
		}
		System.out.println("\nQ:");
		for (int i = 0; i < this.S; i++) {
			System.out.print("\n");
			for (int j = 0; j < this.S; j++) {
				System.out.print(Q[i][j] + "  ");
			}
		}
		System.out.println("\nM:");
		for (int i = 0; i < 4; i++) {
			System.out.print("\n");
			for (int j = 0; j < this.T; j++) {
				System.out.print(M[i][j] + "  ");
			}
		}
		System.out.println("\nP:");
		for (int i = 0; i < this.S; i++) {
			System.out.print("\n");
			for (int j = 0; j < this.T; j++) {
				System.out.print(P[i][j] + "  ");
			}
		}
		System.out.println("\nH:");
		for (int i = 0; i < this.S+2; i++) {
			System.out.print("\n");
			for (int j = 0; j < this.I; j++) {
				System.out.print(H[i][j] + "  ");
			}
		}
		System.out.println("\nE:");
		for (int i = 0; i < this.S+2; i++) {
			System.out.print("\n");
			for (int j = 0; j < this.J; j++) {
				System.out.print(E[i][j] + "  ");
			}
		}
	}
	
	/**
	 * Generates random realistic Q.
	 * @return random Q.
	 */
	public double[][] getQ() {
		double[][] Q = new double[S][S];
		for (int i = 0; i < S; i++) {
			for (int j = 0; j < S; j++) {
				if (rand.nextDouble() < hateSocial)
					Q[i][j] = -0.5 - 0.5 * rand.nextDouble();
			}
		}
		double ceil = socialShare * S;
		int pos = 0;
		while (pos < ceil) {
			int peerSize = 1 + (int) Math.ceil(rand.nextDouble() * (maxPeerSize-1));
			for (int i = pos; i < pos+peerSize; i++) {
				for (int j = pos; j < pos+peerSize; j++) {
					if (i != j)
						Q[i][j] = 0.5 + 0.5 * rand.nextDouble();
				}
			}
			pos += peerSize;
		}
		return Q;
	}
	
	/**
	 * Generates random realistic P.
	 * @return random P.
	 */
	public double[][] getP() {
		double[][] P = new double[S][T];
		for (int i = 0; i < S; i++) {
			for (int j = 0; j < T; j++) {
				double tmp = rand.nextDouble();
				if (tmp < hateTopic)
					P[i][j] = -0.5 - 0.5 * rand.nextDouble();
				else if (tmp > 1-likeTopic)
					P[i][j] = 0.5 + 0.5 * rand.nextDouble();
			}
		}
		return P;
	}
	
	/**
	 * Generates random realistic L.
	 * @return random L.
	 */
	public double[][] getL() {
		double[][] L = new double[S][2];
		for (int i = 0; i < S; i++) {
			L[i][0] = rand.nextDouble();
			L[i][1] = 1 - L[i][0];
		}
		return L;
	}
	
	/**
	 * Generates random realistic instance dimensions.
	 * @return random instance dimensions.
	 */
	public double[][] getSize(double[][] M){
		double[][] Size = new double[5][1];
		Size[0][0] = S;
		Size[2][0] = T;
		Size[3][0] = I;
		Size[4][0] = J;
		double min = S;
		for(int i = 0; i < T; i++) {
			if (M[0][i] < min)
				min = M[0][i];
		}
		if (min > 0)
			Size[1][0] = (int) Math.ceil(S / min);
		else
			Size[1][0] = S;
		return Size;
	}
	
	/**
	 * Generates random realistic M.
	 * @return random M.
	 */
	public double[][] getM() {
		double[][] M = new double[4][T];
		for (int i = 0; i < T; i++) {
			M[0][i] = 3;
			M[1][i] = M[0][i] + 3;
			if (rand.nextDouble() < 0.7/(double)T)
				M[2][i] = 1;
			M[3][i] = M[2][i] + 2;
		}
		return M;
	}
	
	/**
	 * Generates random realistic H.
	 * @return random H.
	 */
	public double[][] getH() {
		double[][] H = new double[S+2][I];
		for (int i = 0; i < I; i++) {
			H[0][i] = (int) Math.floor(2 * rand.nextDouble());
			H[1][i] = 3 + (int) Math.floor(3 * rand.nextDouble());
			double prob = 0.5 + 0.1 * rand.nextDouble();
			int count = 0;
			for (int j = 2; j < S+2; j++) {
				if (rand.nextDouble() < prob) {
					H[j][i] = 1;
					count++;
				}
			}
			for (int k = 0; k < ((double) S) * H[0][i] * ((double)1/3) - count; k++) {
				H[2 + (int) Math.floor(((double)S) * rand.nextDouble())][i] = 1;
			}
		}
		return H;
	}
	
	/**
	 * Generates random realistic E.
	 * @return random E.
	 */
	public double[][] getE() {
		double[][] E = new double[S+2][J];
		for (int i = 0; i < J; i++) {
			E[0][i] = - 0.08 * rand.nextDouble();
			E[1][i] = 0.05 * rand.nextDouble();
			for (int j = 2; j < S+2; j++) {
				E[j][i] = rand.nextDouble();
			}
		}
		return E;
	}
	
}
