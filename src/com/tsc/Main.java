package com.tsc;

public class Main {
	
	static String inDir = "C://Java//Preparse//1.original//";
	static String outDir = "C://Java//Preparse//2.result//";
	static String arcDir = "C://Java//Preparse//3.archive//";
	static String errDir = "C://Java//Preparse//4.error//";
	static String logDir = "C://Java//Preparse//log//";
	static String usage = "C://Java//Preparse//usages.txt";	
		
	// java -jar ./program/Main.java inDir outDir arcDir errDir logDir usages
	public static void main(String[] args) {
				
		Preparse pre = new Preparse();
//		pre.getFoldersFullName();
		
		pre.run();
		
	}
	
	
	
}
