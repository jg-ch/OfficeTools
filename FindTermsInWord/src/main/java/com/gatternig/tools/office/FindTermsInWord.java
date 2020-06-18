/**
 * 
 */
package com.gatternig.tools.office;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
//import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * @author Jochen Gatternig
 *
 */
public class FindTermsInWord {

	String outFile;
	List<String> docs;
	ArrayList<LinkedHashSet<String>> terms;
	HashMap<String, Integer> rowMap;
	HashMap<String, Integer> results;

	/**
	 * 
	 */
	public FindTermsInWord(String[] args) {
		int i = 0;
		long initRT = System.currentTimeMillis();
		BufferedReader lineReader;
		outFile = args[1];

		System.out.println("Start Program...");
		System.out.println("Initializing...");
		terms = new ArrayList<LinkedHashSet<String>>();
		results = new HashMap<String, Integer>();

		// Collect the files to search in
		try (Stream<Path> walk = Files.walk(Paths.get(args[0]))) {
			docs = walk.map(x -> x.toString()).filter(f -> f.endsWith(".docx")).collect(Collectors.toList());
			System.out.println("Documents to process:");
			docs.forEach(System.out::println);
		} catch (IOException e) {
			e.printStackTrace();
		}

		// Reading the stuff to read
		try {
			String searchTerms = args[0] + "\\begriffe.txt";
			rowMap = new HashMap<String, Integer>();
			lineReader = new BufferedReader(new FileReader(searchTerms));
			String line = null;

			while ((line = lineReader.readLine()) != null) {
				List<String> tmp = Arrays.asList(line.split("\\s*,\\s*"));
				tmp.replaceAll(String::toLowerCase);
				rowMap.put(tmp.get(0), i);
				i += 3;
				System.out.println("Topic " + tmp.get(0) + " found");
				LinkedHashSet<String> words = new LinkedHashSet<String>(tmp);
				terms.add(words);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("Initialized after \t" + Long.toString(System.currentTimeMillis() - initRT) + "ms");
	}

	public void execute() {
		long searchDocRT = System.currentTimeMillis();
		for (String doc : docs) {
			long singleDocRT = System.currentTimeMillis();
			String shortDoc = getDocShortName(doc);
			System.out.print("Searching in " + shortDoc);
			System.out.print(".");
			try {
				XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(new FileInputStream(doc)));
				List<XWPFParagraph> paras = xdoc.getParagraphs();
				for (LinkedHashSet<String> set : terms) {
					for (String term : set) {
						for (XWPFParagraph para : paras) {
							String paraText=para.getText().toLowerCase();
							/*List<XWPFRun> runs = para.getRuns();
							for (XWPFRun run : runs) {
								String runText = run.getText(0);*/
								if (paraText != null && paraText.contains(term)) {
									System.out.println("Found " + term);
									addTermToHash(term);
								}
							}
						}
					}
				//}
			} catch (Exception e) {
				e.printStackTrace();
			}
			System.out.println("Done searching in " + shortDoc + " after \t"
					+ Long.toString(System.currentTimeMillis() - singleDocRT) + "ms");
		}
		System.out.println(
				"Done searching terms after \t" + Long.toString(System.currentTimeMillis() - searchDocRT) + "ms");
	}

	private String getDocShortName(String longName) {
		String[] parts = longName.split(Pattern.quote(File.separator));
		return parts[parts.length - 1];
	}

	private void addTermToHash(String term) {
		if (results.containsKey(term)) {
			int i = results.get(term);
			i++;
			results.replace(term, i);
		} else
			results.put(term, 1);
	}

	public void writeExcel() {
		long xlRT = System.currentTimeMillis();
		XSSFWorkbook workbook;
		XSSFSheet sheet;
		System.out.println("Writing data to " + outFile);

		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("Word Analysis Interviews");
		for (LinkedHashSet<String> set : terms) {
			ArrayList<String> lst = new ArrayList<String>(set);
			Row hRow = sheet.createRow(rowMap.get(lst.get(0)));
			Row rRow = sheet.createRow(hRow.getRowNum() + 1);
			int i = 0;
			for (String term : set) {
				Cell cell = hRow.createCell(i);
				cell.setCellValue(term);

				if (results.containsKey(term)) {
					cell = rRow.createCell(i);
					cell.setCellValue(results.get(term));
				}
				i++;
			}
		}

		try {
			workbook.write(new FileOutputStream(outFile));
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println(
				"Done writing " + outFile + " after \t" + Long.toString(System.currentTimeMillis() - xlRT) + "ms");
	}

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		long pgmRT = System.currentTimeMillis();
		if (args.length != 2) {
			System.out.println("Specify Source and Target!");
			System.exit(-1);
		}
		FindTermsInWord ftiw = new FindTermsInWord(args);
		ftiw.execute();
		ftiw.writeExcel();
		System.out.println("Program finished after \t" + Long.toString(System.currentTimeMillis() - pgmRT) + "ms");
	}
}