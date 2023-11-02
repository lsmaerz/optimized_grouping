import java.util.ArrayList;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;
import gurobi.*;

public class StartHere {

	final public static NumberFormat integerNo = new NumberFormat("0");
	final public static WritableCellFormat integerFormat = new WritableCellFormat(integerNo);
	final public static NumberFormat decimalNo = new NumberFormat("0.00");
	final public static WritableCellFormat decimalFormat = new WritableCellFormat(decimalNo);
	
	/**
	 * Choose one of the three options for solving your grouping problem. Your Excel
	 * input file must be formatted as 'MS Excel 1997-2003 sheet' and closed during
	 * execution. Modify the input parameters if needed after reading the
	 * corresponding documentation in this class. Do not modify the source code
	 * elsewhere. See NumericalTests class for fast random instances.
	 * 
	 * @param args (not used)
	 */
	public static void main(String[] args) {
		StartHere.changeLogFile("D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\ConsoleOutput.txt");
		
		//StartHere.frontier("D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\Data\\data_S10_T5.xls", "D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\results_frontier.xls", 300, 0.05);
		//StartHere.multi("NumTest_1234_15_7_2_3", "D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\results_multi.xls", 1000, 0, 1, 0.01);
		//StartHere.single("D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\Data\\data_S10_T5.xls", "D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\results_single.xls", 1000, 0.01, true);
		StartHere.improved("D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\Data\\data_S10_T5.xls", "D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\results_improved.xls", 1000, 0.01, true);
		StartHere.naive("D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\Data\\data_S10_T5.xls", "D:\\Studium\\Semester_9\\Masterseminar_OR\\Public\\Code\\results_naive.xls", 1000, 0.01, true);
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
	 * Solves improved model and queries solution into Excel file
	 * 
	 * @param inputAddress  for user input
	 * @param outputAddress for solution output
	 * @param timeLimit     for finding the solution of the MIP (in s)
	 * @param gap           relative primal-to-dual gap as termination criterion
	 * @param moreInfo      binary information whether or not to compute skill gaps
	 *                      and diversities (only choose TRUE if Excel sheet E is
	 *                      filled)
	 */
	public static void improved(String inputAddress, String outputAddress, int timeLimit, double gap, boolean moreInfo) {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(outputAddress));
			WritableSheet sheet = workbook.createSheet("Results", 0);
			Wrapper.improved(inputAddress, outputAddress, timeLimit, gap, moreInfo, 1, sheet);
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Solves naive model and queries solution into Excel file
	 * 
	 * @param inputAddress  for user input
	 * @param outputAddress for solution output
	 * @param timeLimit     for finding the solution of the MIP (in s)
	 * @param gap           relative primal-to-dual gap as termination criterion
	 * @param moreInfo      binary information whether or not to compute skill gaps
	 *                      and diversities (only choose TRUE if Excel sheet E is
	 *                      filled)              
	 */
	public static void naive(String inputAddress, String outputAddress, int timeLimit, double gap, boolean moreInfo) {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(outputAddress));
			WritableSheet sheet = workbook.createSheet("Results", 0);
			Wrapper.naive(inputAddress, outputAddress, timeLimit, gap, moreInfo, 1, sheet);
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Solves single-objective model and queries solution into Excel file
	 * 
	 * @param inputAddress  for user input
	 * @param outputAddress for solution output
	 * @param timeLimit     for finding the solution of the MIP (in s)
	 * @param gap           relative primal-to-dual gap as termination criterion
	 * @param moreInfo      binary information whether or not to compute skill gaps
	 *                      and diversities (only choose TRUE if Excel sheets H and
	 *                      E are filled)
	 */
	public static void single(String inputAddress, String outputAddress, int timeLimit, double gap, boolean moreInfo) {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(outputAddress));
			WritableSheet sheet = workbook.createSheet("Results", 0);
			Wrapper.single(inputAddress, outputAddress, timeLimit, gap, moreInfo, 1, sheet);
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Solves multi-criteria model and queries solution into Excel file
	 * 
	 * @param inputAddress  for user input
	 * @param outputAddress for solution output
	 * @param timeLimit     for finding the solution of the multi-objective MIP (in s)
	 * @param prio          priority of the skill distribution objective, preference
	 *                      term has prio 0, hence -1,0,1 cover all
	 *                      hierarchical/blended cases
	 * @param weightOrTol   weight in blended case (prio = 0) and tolerance in
	 *                      hierarchical case (prio != 0)
	 * @param pos			position of output row  
	 * @param gap           relative primal-to-dual gap as termination criterion
	 */
	public static void multi(String inputAddress, String outputAddress, int timeLimit, int prio, double weightOrTol, double gap) {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(outputAddress));
			WritableSheet sheet = workbook.createSheet("Results", 0);
			Wrapper.multi(inputAddress, outputAddress, timeLimit, prio, weightOrTol, gap, 1, sheet);
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Solves multi-criteria model for a collection of priorities and
	 * weights/tolerances, then queries solutions into Excel file to approximate efficient frontier
	 * 
	 * @param inputAddress  for user input
	 * @param outputAddress for solution output
	 * @param timeLimit     for comuting the whole frontier (in s)
	 * @param gap           relative primal-to-dual gap as termination criterion
	 */
	public static void frontier(String inputAddress, String outputAddress, int timeLimit, double gap) {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(outputAddress));
			WritableSheet sheet = workbook.createSheet("Results", 0);
			Wrapper.frontier(inputAddress, outputAddress, timeLimit, gap, sheet);
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}