import java.io.FileInputStream;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.*;

import javax.imageio.ImageIO;
import javax.swing.JFileChooser;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class PapGen 
{
	/*
	 	EleId:
	 	0 - Header
	 	1 - Part
	 	2 - Question
	 	3 - Image
	 	4 - OR
	*/
	
	static ArrayList<PaperElement> Paper = new ArrayList<PaperElement>();
	static List<XWPFPictureData> Pics;
	static List<XWPFPictureData> DefaultPics;
	static String outputPath = "C:\\Users\\Lenovo\\Desktop\\";
	static String outfile = "Final Paper"; 
	static String defaultPicsPath = "C:\\Work Material\\PapGen\\Image Container.docx";
	
	public static void main(String args[]) throws FileNotFoundException, IOException, InvalidFormatException
	{
		
		// ch = choice, m = marks, mod = module
		// Place Question Paper Header and Question Banks Folder Path Here
		
		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
		
		XWPFDocument PicDoc= new XWPFDocument(new FileInputStream(defaultPicsPath));
		DefaultPics = PicDoc.getAllPictures();
		
		Paper.add(new PaperElement(0, ""));
		
		System.out.println("\n\nWelcome To PapGen\n"
				+ "Enter the required details to generate a Question Paper.\n");
		
		int ch=0;
		
		System.out.println("\n Press 0 for Default Paper Generation\n Press 1 for Manual Paper Generation\n");
		ch = Integer.parseInt(br.readLine());
		
		if(ch == 0)
			defaultMode();
		else
			manualMode();
	}
	
	static void manualMode() throws FileNotFoundException, IOException, InvalidFormatException
	{
		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
		int ch = 0;
		
		do
		{
			System.out.println("\n1. Insert Part\n"
					+ "2. Insert Question\n"
					+ "3. Undo Previous Action\n"
					+ "4. Insert OR\n"
					+ "5. Generate Paper\n"
					+ "6. Exit\n\n"
					+ "Enter your choice:\n");
			
			ch = Integer.parseInt(br.readLine());
			
			switch(ch)
			{
				case 1: 
					System.out.println("\nEnter Part Name:");
					Paper.add(new PaperElement(1, "Part " + br.readLine()));
					break;
				case 2: 
					System.out.println("\nEnter Question Number:");
					String num = br.readLine();
					
					System.out.println("\nEnter required Marks:");
					int m = Integer.parseInt(br.readLine());
					
					System.out.println("\nChoose the Question Bank from which the question is to be picked:"
							+ "\n(Press any key to open the Browse Window on the Desktop)");
					String dummy = br.readLine();
					
					String fileContents = readFile();
					if(fileContents.equals("-1"))
						System.out.println("File Not Read!");
					
					String SL = "";	
					String quesString = "";
					
					do
					{
						quesString = quesPicker(fileContents,m);
						if(quesString.equals(""))
							break;
						int end = quesString.charAt(0) - '0';
						SL = quesString.substring(1, end + 1);
						quesString = quesString.substring(SL.length() + 1);
					}while(duplicate(quesString));
					
					if(quesString.equals(""))
					{	
						System.out.println("Question Bank not formatted correctly, or No such question.");
						break;
					}
					
					int ind = Integer.parseInt(SL);
					
					Paper.add(new PaperElement(2, num + " " + quesString));
					XWPFPictureData pic = Pics.get(ind - 1);
					Paper.add(new PaperElement(3, pic));
					break;
				case 3:
					Paper.remove(Paper.size() - 1);
					break;
				case 4:
					Paper.add(new PaperElement(4, "OR"));
					break;
				case 5:
					writeFile();
					System.out.println("\nA Word Document of the Question Paper has been generated to your Output Path.\n");
					break;
				case 6:
					System.exit(0);
					break;
				default:
			}
			
		}while(ch != 6);
	}
	
	static void defaultMode() throws FileNotFoundException, IOException, InvalidFormatException
	{
		/*
		 ch:
		 	1 - Insert Part
			2 - Insert Question
			3 - Insert OR
		 	4 - Generate Paper
		 	
		 qCounter - Question Counter: 1 - 6
		 */
		
		char partName = 'A';
		int ch = 0, op = 0, qCounter = 1;
		int options[] = {1, 2, 3, 2, 2, 3, 2, 1, 2, 3, 2, 4};
		int A[] = {2, 3, 2, 3};
		int B[] = {2, 1};
		
		while(op <= 11)
		{
			ch = options[op];
			
			switch(ch)
			{
				case 1: 
					Paper.add(new PaperElement(1, "Part " + partName));
					partName += 1;
					op++;
					break;
					
				case 2:
					if(partName == 'B') // Because 'A' incremented to 'B'
					{
						int success = defaultQuestionSetter(qCounter, A[qCounter - 1], 20);
						if(success == 0)
						{
							int temp = A[qCounter - 1] - 1;
							if(temp != 1)
							{
								success = defaultQuestionSetter(qCounter, temp, 20);
								if(success == 0)
								{
									op--;
									System.out.println("Question Bank does not contain enough variety of questions.");
								}
								else
									qCounter++;
							}
							else
							{
								op--;
								System.out.println("Question Bank does not contain enough variety of questions.");
							}                                                                                                                     	op--;
                                                                                                                                                	System.out.println("Question Bank does not contain enough variety of questions.");
                         }
						 else
							 qCounter++;
					}
					else
					{
						int success = defaultQuestionSetter(qCounter, B[qCounter - 5], 10);
						if(success == 0)
						{
							int temp = B[qCounter - 5] - 1;
							if(temp != 0)
							{
								success = defaultQuestionSetter(qCounter, temp, 10);
								if(success == 0)
									System.out.println("Question Bank does not contain enough variety of questions.");
								else
									qCounter++;
							}
							else
								System.out.println("Question Bank does not contain enough variety of questions.");
						}
						else
							qCounter++;
					}
					op++;
					break;
				case 3:
					Paper.add(new PaperElement(4, "OR"));
					op++;
					break;
				case 4:
					writeFile();
					System.out.println("\nA Word Document of the Question Paper has been generated to your Output Path.\n");
					op++;
					break;
			}
		}
	}
	
	static String readFile() throws FileNotFoundException, IOException
	{
		JFileChooser file= new JFileChooser();
		int fileRetValue= file.showOpenDialog(null);
		
		if(fileRetValue == JFileChooser.APPROVE_OPTION) 
		{
			XWPFDocument document= new XWPFDocument(new FileInputStream(file.getSelectedFile()));
			XWPFWordExtractor extract= new XWPFWordExtractor(document);
			String fc = extract.getText();
			extract.close();
			Pics = document.getAllPictures();
			return fc;
		}
		return "-1";
	}
	
	static String quesPicker(String fileContents, int m)
	{
		StringTokenizer st = new StringTokenizer(fileContents,"\n");
		
		ArrayList<String> quesList = new ArrayList<String>();
		String line = "";
		
		while(st.hasMoreTokens())
		{
			if(line.startsWith("_Marks_"))
			{
				String x; // marks
				int lu = line.lastIndexOf("_"); // Last index of Underscore
				String s = line.substring(lu + 1); // SL NO
				
				if(m > 9)
					x = line.substring(7, 9);
				else
					x = line.substring(7, 8);
				if(x.equals(String.valueOf(m)))
				{
					String quesString = "";
					line = st.nextToken().trim();
					while(!line.startsWith("_Marks_"))
					{
						quesString += line + "\n";
						if(st.hasMoreTokens())
							line = st.nextToken().trim();
						else
							break;
					}
					quesList.add(s.length()+s+quesString);
				}
				else
					line = st.nextToken().trim();
			}
			else
				line = st.nextToken().trim();
		}
		int max = quesList.size() - 1;
		if(max == -1)
			return "";
		int rand = (int)Math.round(Math.random()*max);
		return quesList.get(rand);
	}
	
	static boolean duplicate(String quesString)
	{
		for(int i=0; i<Paper.size(); i++)
		{
			if(Paper.get(i).getID() == 2)
			{
				if(Paper.get(i).getContent().endsWith(quesString))
					return true;
			}
		}
		return false;
	}
	
	static void writeFile() throws IOException, InvalidFormatException
	{
		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
		
		System.out.println("\nEnter a unique name for the file to be generated:");
		outfile = br.readLine();
		
		XWPFDocument doc = new XWPFDocument();
		FileOutputStream out = new FileOutputStream(new File(outputPath + outfile + ".docx"));
		
		XWPFParagraph para;
		XWPFRun run;
		
		for(int i=0; i<Paper.size(); i++)
		{
			int ID = Paper.get(i).getID();
			String ques;
			
			switch(ID)
			{
				case 0:
					
					para = doc.createParagraph();
					para.setAlignment(ParagraphAlignment.RIGHT);
					run = para.createRun();
					run.setBold(true);
					run.setFontSize(10);
					run.setFontFamily("Calibri (Body)");
					run.setCapitalized(true);
					run.setText("USN: __ __ __ __ __ __ __ __ __ __");
					run.addBreak();
					
					para = doc.createParagraph();
					para.setAlignment(ParagraphAlignment.CENTER);
					run = para.createRun();
					run.setBold(true);
					run.setFontSize(10);
					run.setFontFamily("Calibri (Body)");
					run.setCapitalized(true);
					run.setText("DAYANANDA SAGAR COLLEGE OF ENGINNERING");
					run.addCarriageReturn();
					run.setText("DEPARTMENT OF COMPUTER SCEIENCE & ENGINEERING");
					run.addCarriageReturn();
					run.setText("__ INTERNAL TEST");
					run.addBreak();
					
					para = doc.createParagraph();
					para.setAlignment(ParagraphAlignment.LEFT);
					run = para.createRun();
					run.setBold(true);
					run.setFontSize(9);
					run.setFontFamily("Calibri (Body)");
					run.setCapitalized(true);
					run.setText("SUB :                                                                                                                                                         CODE :                       ");
					run.setText("TIME :                                                                                                                                                      MAX. MARKS :50    ");
					run.setText("DATE :                                                                                                                                                      SEC :                         ");
					run.addBreak();
					
					para = doc.createParagraph();
					para.setAlignment(ParagraphAlignment.CENTER);
					run = para.createRun();
					run.setBold(true);
					run.setItalic(true);
					run.setUnderline(UnderlinePatterns.SINGLE);
					run.setFontSize(9);
					run.setFontFamily("Calibri (Body)");
					run.setText("Note: Answer any two full questions from Part A and ");
					run.setText("one full question from Part B.");
					
					para = doc.createParagraph();
					para.setAlignment(ParagraphAlignment.CENTER);
					run = para.createRun();
					run.setBold(false);
					run.setItalic(true);
					run.setFontSize(9);
					run.setFontFamily("Calibri (Body)");
					run.setText("Support your answers with diagrams/structures wherever necessary.");
					
					para = doc.createParagraph();
					para.setAlignment(ParagraphAlignment.CENTER);
					run = para.createRun();
					run.setBold(false);
					run.setItalic(true);
					run.setFontSize(9);
					run.setFontFamily("Calibri (Body)");
					run.setText("_____________________________________________________________________________________________________");
					run.addBreak();
					
					break;
				
				case 1:
					para = doc.createParagraph();
					para.setAlignment(ParagraphAlignment.CENTER);
					run = para.createRun();
					
					ques = Paper.get(i).getContent();
					run.setFontSize(9);
					run.setBold(true);
					run.setFontFamily("Calibri (Body)");
					run.setCapitalized(true);
					run.setUnderline(UnderlinePatterns.SINGLE);
					run.setText(ques);
					run.addBreak();
					break;
					
				case 2:
					para = doc.createParagraph();
					para.setAlignment(ParagraphAlignment.LEFT);
					run = para.createRun();
					run.setFontFamily("Calibri (Body)");
					run.setFontSize(9);
					run.setFontFamily("Times New Roman");
					
					ques = Paper.get(i).getContent();
					StringTokenizer QuesSt = new StringTokenizer(ques, "\n");
					while(QuesSt.hasMoreTokens())
					{
						run.setText(QuesSt.nextToken());
					}
					break;
					
				case 3:
					XWPFPictureData pic= Paper.get(i).getImage();
					if(defaultChecker(pic))
					{
						para = doc.createParagraph();
						para.setAlignment(ParagraphAlignment.CENTER);
						run = para.createRun();
					
						run.addPicture(new ByteArrayInputStream(pic.getData()), pic.getPictureType(), "Final Paper.docx", Units.toEMU(100), Units.toEMU(100));
						run.addBreak();
					}
					break;
					
				case 4:
					para = doc.createParagraph();
					para.setAlignment(ParagraphAlignment.CENTER);
					run = para.createRun();
					
					ques = Paper.get(i).getContent();
					run.setBold(true);
					run.setUnderline(UnderlinePatterns.SINGLE);
					run.setText(ques);
					run.addBreak();
					break;
			}
		}
		doc.write(out);
		out.close();
		doc.close();
	}
	
	static boolean defaultChecker(XWPFPictureData pic)
	{
		byte bytePic[] = pic.getData();
		Iterator<XWPFPictureData> iterator = DefaultPics.iterator();
		
		while(iterator.hasNext())
		{
		   byte byteSample[] = iterator.next().getData();
		   if(Arrays.equals(bytePic, byteSample))
			   return false;
		}
		return true;
	}
	
	static int defaultQuestionSetter(int qNo, int no_of_ques, int total) throws IOException, FileNotFoundException
	{
		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
		int m = 0, iter = 1, initial = Paper.size();
		int marks[] = {5, 6, 7, 8, 10, 10, 12};
		
		System.out.println("\nChoose the Question Bank from which question " + qNo + " is to be picked:"
				+ "\n(Press any key to open the Browse Window on the Desktop)");
		String dummy = br.readLine();
		
		String fileContents = readFile();
		if(fileContents.equals("-1"))
		{
			System.out.println("File Not Read!");
			return 0;
		}
		
		String SL = "";	
		String quesString = "";
		char letter = 'a';
		int marksBreakUp[] = {0, 0, 0};
		int check = 0;
		
		outer: for(int i=0; (i < no_of_ques) && (iter < 300); i++)
		{
			String num = String.valueOf(qNo) + ". ";
			
			if(no_of_ques > 1)
			{
				letter = 'a';
				letter += i;
				num += "(" + letter + ") ";
			}
			
			if(no_of_ques == 1)
				m = total;
			else if(no_of_ques == 2)
			{
				while(check != total)
				{
					marksBreakUp[0] = marks[(int)Math.round(Math.random()*6)];
					marksBreakUp[1] = marks[(int)Math.round(Math.random()*6)];
					check = marksBreakUp[0] + marksBreakUp[1];
				}
				if(letter == 'a')
					m = marksBreakUp[0];
				else
					m = marksBreakUp[1];
			}
			else
			{
				while(check != total)
				{
					marksBreakUp[0] = marks[(int)Math.round(Math.random()*6)];
					marksBreakUp[1] = marks[(int)Math.round(Math.random()*6)];
					marksBreakUp[2] = marks[(int)Math.round(Math.random()*6)];
					check = marksBreakUp[0] + marksBreakUp[1] + marksBreakUp[2];
				}
				if(letter == 'a')
					m = marksBreakUp[0];
				if(letter == 'b')
					m = marksBreakUp[1];
				else
					m = marksBreakUp[2];
			}
			
			do
			{
				quesString = quesPicker(fileContents,m);
				if(quesString.equals(""))
					break;
				int end = quesString.charAt(0) - '0';
				SL = quesString.substring(1, end + 1);
				quesString = quesString.substring(SL.length() + 1);
			}while(duplicate(quesString));
			
			if(quesString.equals("")) // failed search case
			{
				i = 0;
				iter++;
				check = 0;
				
				if((Paper.size() - initial) == 2)
					Paper.remove(Paper.size() - 1);
				else if((Paper.size() - initial) == 4)
				{
					Paper.remove(Paper.size() - 1);
					Paper.remove(Paper.size() - 1);
				}
				
				continue outer;
			}
			
			int ind = Integer.parseInt(SL);
			
			Paper.add(new PaperElement(2, num + " " + quesString));
			XWPFPictureData pic = Pics.get(ind - 1);
			Paper.add(new PaperElement(3, pic));
		}
		
		if((Paper.size() - initial) == (no_of_ques*2))
			return 1;
		return 0;
	}
}