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

public class Wrapper {

	final public static NumberFormat integerNo = new NumberFormat("0");
	final public static WritableCellFormat integerFormat = new WritableCellFormat(integerNo);
	final public static NumberFormat decimalNo = new NumberFormat("0.00");
	final public static WritableCellFormat decimalFormat = new WritableCellFormat(decimalNo);
	
	public static void main(String[] args) {}

	/**
	 * Solves improved model and queries solution into Excel file
	 * 
	 * @param inputAddress  for user input
	 * @param outputAddress for solution output
	 * @param timeLimit     for finding the solution of the MIP
	 * @param gap           relative primal-to-dual gap as termination criterion
	 * @param moreInfo      binary information whether or not to compute skill gaps
	 *                      and diversities (only choose TRUE if Excel sheet E is
	 *                      filled)
	 * @param pos			current row to be written in
	 * @param sheet			current sheet to be written on
	 * @return				current sheet written on
	 */
	public static void improved(String inputAddress, String outputAddress, int timeLimit, double gap, boolean moreInfo, int pos, WritableSheet sheet) {
		long start, end;
		try {
			if (moreInfo) {
				Label label = new Label(0, 0, "ObjVal");
				sheet.addCell(label);
				label = new Label(1, 0, "SkillVal");
				sheet.addCell(label);
				label = new Label(2, 0, "AvgSocialSat");
				sheet.addCell(label);
				label = new Label(3, 0, "RunTime (in ms)");
				sheet.addCell(label);
				label = new Label(4, 0, "Assignment");
				sheet.addCell(label);
				label = new Label(5, 0, "SkillDistr (Gaps% & Diversity)");
				sheet.addCell(label);
				start = System.nanoTime();
				StandardImproved.execute(sheet, timeLimit, gap, inputAddress, moreInfo, pos);
				end = System.nanoTime();
				Number num = new Number(3, pos, (end - start) / 1000000, integerFormat);
				sheet.addCell(num);
			} else {
				Label label = new Label(0, 0, "ObjVal");
				sheet.addCell(label);
				label = new Label(1, 0, "AvgSocialSat");
				sheet.addCell(label);
				label = new Label(2, 0, "RunTime (in ms)");
				sheet.addCell(label);
				label = new Label(3, 0, "Assignment");
				sheet.addCell(label);
				start = System.nanoTime();
				StandardImproved.execute(sheet, timeLimit, gap, inputAddress, moreInfo, pos);
				end = System.nanoTime();
				Number num = new Number(2, pos, (end - start) / 1000000, integerFormat);
				sheet.addCell(num);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Solves naive model and queries solution into Excel file
	 * 
	 * @param inputAddress  for user input
	 * @param outputAddress for solution output
	 * @param timeLimit     for finding the solution of the MIP
	 * @param gap           relative primal-to-dual gap as termination criterion
	 * @param moreInfo      binary information whether or not to compute skill gaps
	 *                      and diversities (only choose TRUE if Excel sheet E is
	 *                      filled)
	 * @param pos			current row to be written in
	 * @param sheet			current sheet to be written on
	 * @return				current sheet written on
	 */
	public static void naive(String inputAddress, String outputAddress, int timeLimit, double gap, boolean moreInfo, int pos, WritableSheet sheet) {
		long start, end;
		try {
			if (moreInfo) {
				Label label = new Label(0, 0, "ObjVal");
				sheet.addCell(label);
				label = new Label(1, 0, "SkillVal");
				sheet.addCell(label);
				label = new Label(2, 0, "AvgSocialSat");
				sheet.addCell(label);
				label = new Label(3, 0, "RunTime (in ms)");
				sheet.addCell(label);
				label = new Label(4, 0, "Assignment");
				sheet.addCell(label);
				label = new Label(5, 0, "SkillDistr (Gaps% & Diversity)");
				sheet.addCell(label);
				start = System.nanoTime();
				StandardNaive.execute(sheet, timeLimit, gap, inputAddress, moreInfo, pos);
				end = System.nanoTime();
				Number num = new Number(3, pos, (end - start) / 1000000, integerFormat);
				sheet.addCell(num);
			} else {
				Label label = new Label(0, 0, "ObjVal");
				sheet.addCell(label);
				label = new Label(1, 0, "AvgSocialSat");
				sheet.addCell(label);
				label = new Label(2, 0, "RunTime (in ms)");
				sheet.addCell(label);
				label = new Label(3, 0, "Assignment");
				sheet.addCell(label);
				start = System.nanoTime();
				StandardNaive.execute(sheet, timeLimit, gap, inputAddress, moreInfo, pos);
				end = System.nanoTime();
				Number num = new Number(2, pos, (end - start) / 1000000, integerFormat);
				sheet.addCell(num);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Solves ExtendedForTopicAssignment model and queries solution into Excel file
	 * 
	 * @param inputAddress  for user input
	 * @param outputAddress for solution output
	 * @param timeLimit     for finding the solution of the MIP
	 * @param gap           relative primal-to-dual gap as termination criterion
	 * @param moreInfo      binary information whether or not to compute skill gaps
	 *                      and diversities (only choose TRUE if Excel sheet E is
	 *                      filled)
	 * @param pos			current row to be written in
	 * @param sheet			current sheet to be written on
	 * @return				current sheet written on
	 */
	public static void single(String inputAddress, String outputAddress, int timeLimit, double gap, boolean moreInfo, int pos, WritableSheet sheet) {
		long start, end;
		try {
			if (moreInfo) {
				Label label = new Label(0, 0, "PrefVal");
				sheet.addCell(label);
				label = new Label(1, 0, "SkillVal");
				sheet.addCell(label);
				label = new Label(2, 0, "AvgSocialSat");
				sheet.addCell(label);
				label = new Label(3, 0, "AvgTopicSat");
				sheet.addCell(label);
				label = new Label(4, 0, "RunTime (in ms)");
				sheet.addCell(label);
				label = new Label(5, 0, "Assignment");
				sheet.addCell(label);
				label = new Label(6, 0, "SkillDistr (Gaps% & Diversity)");
				sheet.addCell(label);
				start = System.nanoTime();
				ExtendedForTopicAssignment.execute(sheet, timeLimit, gap, inputAddress, moreInfo, pos);
				end = System.nanoTime();
				Number num = new Number(4, pos, (end - start) / 1000000, integerFormat);
				sheet.addCell(num);
			} else {
				Label label = new Label(0, 0, "PrefVal");
				sheet.addCell(label);
				label = new Label(1, 0, "AvgSocialSat");
				sheet.addCell(label);
				label = new Label(2, 0, "AvgTopicSat");
				sheet.addCell(label);
				label = new Label(3, 0, "RunTime (in ms)");
				sheet.addCell(label);
				label = new Label(4, 0, "Assignment");
				sheet.addCell(label);
				start = System.nanoTime();
				ExtendedForTopicAssignment.execute(sheet, timeLimit, gap, inputAddress, moreInfo, pos);
				end = System.nanoTime();
				Number num = new Number(3, pos, (end - start) / 1000000, integerFormat);
				sheet.addCell(num);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Solves multi-criteria model and queries solution into Excel file
	 * 
	 * @param inputAddress  for user input
	 * @param outputAddress for solution output
	 * @param timeLimit     for finding the solution of the MIP
	 * @param prio          priority of the skill distribution objective, preference
	 *                      term has prio 0, hence -1,0,1 cover all
	 *                      hierarchical/blended cases
	 * @param weightOrTol   weight in blended case (prio = 0) and tolerance in
	 *                      hierarchical case (prio != 0)
	 * @param pos			position of output row
	 * @param sheet			current sheet to be written on  
	 * @param gap           relative primal-to-dual gap as termination criterion
	 * @return				current sheet written on
	 */
	public static void multi(String inputAddress, String outputAddress, int timeLimit, int prio, double weightOrTol, double gap, int pos, WritableSheet sheet) {
		long start, end;
		try {
			Label label = new Label(0, 0, "Prio");
			sheet.addCell(label);
			label = new Label(1, 0, "Tol & Weight");
			sheet.addCell(label);
			label = new Label(2, 0, "PrefVal");
			sheet.addCell(label);
			label = new Label(3, 0, "SkillVal");
			sheet.addCell(label);
			label = new Label(4, 0, "AvgSocialSat");
			sheet.addCell(label);
			label = new Label(5, 0, "AvgTopicSat");
			sheet.addCell(label);
			label = new Label(6, 0, "RunTime (in ms)");
			sheet.addCell(label);
			label = new Label(7, 0, "Assignment");
			sheet.addCell(label);
			label = new Label(8, 0, "MIP Gap");
			sheet.addCell(label);
			label = new Label(9, 0, "SkillDistr (Gaps% & Diversity)");
			sheet.addCell(label);
			Number num = new Number(0, pos, prio, integerFormat);
			sheet.addCell(num);
			num = new Number(1, pos, weightOrTol, decimalFormat);
			sheet.addCell(num);
			start = System.nanoTime();
			if (prio == 0)
				ExtendedForMultiCriteria.execute(sheet, prio, 0, weightOrTol, pos, timeLimit, gap, inputAddress);
			else
				ExtendedForMultiCriteria.execute(sheet, prio, weightOrTol, 1, pos, timeLimit, gap, inputAddress);
			end = System.nanoTime();
			num = new Number(6, pos, (end - start) / 1000000, integerFormat);
			sheet.addCell(num);
		} catch (Exception e) {
			System.out.println("HIER vor");
			e.printStackTrace();
			System.out.println("HIER nach");
		}
	}

	/**
	 * Solves multi-criteria model for a collection of priorities and
	 * weights/tolerances, then queries solutions into Excel file to approximate efficient frontier
	 * 
	 * @param inputAddress  for user input
	 * @param outputAddress for solution output
	 * @param timeLimit     for finding the solution of the MIP
	 * @param gap           relative primal-to-dual gap as termination criterion
	 */
	public static void frontier(String inputAddress, String outputAddress, int timeLimit, double gap, WritableSheet sheet) {
		int pos = 1;
		double weight, tol;
		long start, end;
		try {
			Label label = new Label(0, 0, "Prio");
			sheet.addCell(label);
			label = new Label(1, 0, "Tol & Weight");
			sheet.addCell(label);
			label = new Label(2, 0, "PrefVal");
			sheet.addCell(label);
			label = new Label(3, 0, "SkillVal");
			sheet.addCell(label);
			label = new Label(4, 0, "AvgSocialSat");
			sheet.addCell(label);
			label = new Label(5, 0, "AvgTopicSat");
			sheet.addCell(label);
			label = new Label(6, 0, "RunTime (in ms)");
			sheet.addCell(label);
			label = new Label(7, 0, "Assignment");
			sheet.addCell(label);
			label = new Label(8, 0, "SkillDistr (Gaps% & Diversity)");
			sheet.addCell(label);
			for (int prio = -1; prio < 2; prio++) {
				if (prio == 0) {
					weight = 1;
					Number num = new Number(0, pos, prio, integerFormat);
					sheet.addCell(num);
					num = new Number(1, pos, weight, decimalFormat);
					sheet.addCell(num);
					start = System.nanoTime();
					ExtendedForMultiCriteria.execute(sheet, prio, 0, weight, pos, (int) Math.floor(((double) timeLimit)/17), gap, inputAddress);
					end = System.nanoTime();
					num = new Number(6, pos, (end - start) / 1000000, integerFormat);
					sheet.addCell(num);
					pos++;
					/**
					weight = 0;
					num = new Number(0, pos, prio, integerFormat);
					sheet.addCell(num);
					num = new Number(1, pos, weight, decimalFormat);
					sheet.addCell(num);
					start = System.nanoTime();
					sheet = ExtendedForMultiCriteriaHierarchical.execute(sheet, prio, 0, weight, pos, timeLimit, gap,
							inputAddress);
					end = System.nanoTime();
					num = new Number(6, pos, (end - start) / 1000000, integerFormat);
					sheet.addCell(num);
					pos++;
					**/
					for (int w = 1; w < 4; w++) {
						weight = Math.pow(2, w);
						num = new Number(0, pos, prio, integerFormat);
						sheet.addCell(num);
						num = new Number(1, pos, weight, decimalFormat);
						sheet.addCell(num);
						start = System.nanoTime();
						ExtendedForMultiCriteria.execute(sheet, prio, 0, weight, pos, (int) Math.floor(((double) timeLimit)/17), gap, inputAddress);
						end = System.nanoTime();
						num = new Number(6, pos, (end - start) / 1000000, integerFormat);
						sheet.addCell(num);
						pos++;
						weight = 1 / weight;
						num = new Number(0, pos, prio, integerFormat);
						sheet.addCell(num);
						num = new Number(1, pos, weight, decimalFormat);
						sheet.addCell(num);
						start = System.nanoTime();
						ExtendedForMultiCriteria.execute(sheet, prio, 0, weight, pos, (int) Math.floor(((double) timeLimit)/17), gap, inputAddress);
						end = System.nanoTime();
						num = new Number(6, pos, (end - start) / 1000000, integerFormat);
						sheet.addCell(num);
						pos++;
					}

				} else {
					for (int t = 5; t < 10; t++) {
						tol = t;
						tol = tol / 10;
						Number num = new Number(0, pos, prio, integerFormat);
						sheet.addCell(num);
						num = new Number(1, pos, tol, decimalFormat);
						sheet.addCell(num);
						start = System.nanoTime();
						ExtendedForMultiCriteria.execute(sheet, prio, tol, 1, pos, (int) Math.floor(((double) timeLimit)/17), gap, inputAddress);
						end = System.nanoTime();
						num = new Number(6, pos, (end - start) / 1000000, integerFormat);
						sheet.addCell(num);
						pos++;
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}