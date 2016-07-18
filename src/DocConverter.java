 import java.util.StringTokenizer;

import com.aspose.cells.TextAlignmentType;
import com.aspose.words.*;

 
public class DocConverter {
	Document doc;			//the .doc file
	Paragraph par;			//the active paragragh
	Run run;				//the object that handles text
	DocumentBuilder builder ;
	
	static public boolean isBold;					//checks if bold tag is active , [b]=true [/b]=false
	public	boolean isItalic;				//checks if italic is tag active , [i]=true [/i]=false		
	boolean isLeft;					//checks if left tag is active , [l]=true [/l]=false	
	boolean isCenter;				//checks if center tag is active , [c]=true [/c]=false
	boolean isRight;				//checks if right tag is active , [r]=true [/r]=false
	boolean ChartCreation;			//true when the next lines are gonna be about chart, false when each chart is finished
	//constractor of the new Document
	CChart customChart;				//temp new custum Chart
	Font font;						//font variable
	 ParagraphFormat paragraphFormat ;	

	 
	 //table components
	 boolean TableCreation;
	 boolean headers;
	 boolean data;
	 int columns;
	 int tableID;
	 Table table;
	private String tableStyle;
	 
	 
	//constractor
	public DocConverter() throws Exception{
		doc = new Document();										//new .doc
													
		isItalic = isBold = isLeft = isCenter = isRight = ChartCreation= false;		//in the start, all tags are closed
		 builder = new DocumentBuilder(doc);			//the builder insert the text into the doc
		 font= builder.getFont();
		  paragraphFormat = builder.getParagraphFormat();		//creating a new paragraph to start
		  paragraphFormat.setAlignment(TextAlignmentType.LEFT);
		  run=new Run(doc);
		  headers=data=TableCreation=false;	//closing all tags by default
		  tableID=columns=0;		//set the current number of tables and columns at zero
	}
	
	//each line goes into that method, and depend on the str(line of Main method)
	//parametr, either handles a tag or adds text
	public void search(String str) throws Exception{

		String temp = str;				//keeping a backup of the line
		str=str.replaceAll(" ", "");	//remove all spaces ( helps in table creation )
		
		//if a new chart starts
		if(str.equalsIgnoreCase("[excel-xx]") || str.equalsIgnoreCase("[excel-chart]") || str.equalsIgnoreCase("[excel-scatter]") 
				|| str.equalsIgnoreCase("[excel-columns]") ){
			
			customChart = new CChart(str);	//make a new CChart object with the given style
			ChartCreation=true;				//set that the next lines will be about charts and will go in CChart.edit
		
		//if chart is finished
		}else if(str.equalsIgnoreCase("[/excel-xx]") || str.equalsIgnoreCase("[/excel-chart]") || str.equalsIgnoreCase("[/excel-scatter]")
				|| str.equalsIgnoreCase("[/excel-columns]")){
			
			customChart.save();	//save the chart by making the image of the chart
			ChartCreation=false;//stop adding parametrs into the chart
			builder.insertImage( "temp.png");	//add the image of the chart
		//if chart creation is active
		}else if(ChartCreation){
			customChart.edit(temp);	//sent the tag into the class CChart
		//if a new table starts
		}else if(str.equalsIgnoreCase("[table]")){
			
			table = (Table)doc.getChild(NodeType.TABLE, tableID, true); //set the new table in the next available ID (tableID)
			builder.startTable();		//make a new table
			TableCreation=true;			//active the table creation
		
		//if table ends
		}else if(str.equalsIgnoreCase("[/table]")){
			
			builder.endTable();		//finish the table
			TableCreation=false;	//diactive the tag
			setTableStyle(tableStyle);	//set the style of the table from this method
			tableID++;				//give the next ID
		//font size tag	
		}else if(str.startsWith("[s=")){
			
			//try to parse integer
			  try {
		        
				  font.setSize(Integer.parseInt(str.substring(3,str.length()-1)));
		        //if failed, parse double   
		        } catch (NumberFormatException e) {
		       
		        	font.setSize(Double.parseDouble(str.substring(3,str.length()-1)));
		        }
			
		//color of font
		}else if(str.startsWith("[color=")){
			
			font.setColor(findColor(str.substring(7,str.length()-1)));  //the color is returned from findColor method
		//font of the paragraph	
		}else if(str.startsWith("[font=")){
			
			font=run.getFont();
			font.setName(str.substring(7,str.length()-1));
		
		//new paragraph
		}else if(str.equalsIgnoreCase("[p]")){						
			
			builder.writeln ();
			
		
		}else if (str.equalsIgnoreCase("[c]")){					//if line was [c]
																//activate the center tag
			newTag(TextAlignmentType.CENTER);										//and aplly it
				 
		
		}else if(str.equalsIgnoreCase("[/c]")){					//if line was [/c]
																//activate the center tag
			newTag(TextAlignmentType.LEFT);										//and aplly it
				
		
		}else if (str.equalsIgnoreCase("[r]")){					//if line was [r]
			
			newTag(TextAlignmentType.RIGHT);			 
			
		}else if(str.equalsIgnoreCase("[/r]")){					//if line was [/r]
												//deactivate the right tag
			newTag(TextAlignmentType.RIGHT);										//and aplly it
				
		
		}else if (str.equalsIgnoreCase("[l]")){					//if line was [l]
											//activate the left tag
			newTag(TextAlignmentType.LEFT);										//and aplly it
						 
		
		}else if(str.equalsIgnoreCase("[/l]")){					//if line was [/l]
																//deactivate the left tag
			newTag(TextAlignmentType.LEFT);										//and aplly it

		}else if (str.equalsIgnoreCase("[b]")){					//if line was [b]
		
			isBold=true;										//activate the bold tag
			font.setBold(true);									//and aplly it
		
		}else if (str.equalsIgnoreCase("[/b]")){				//if line was [/b]
			isBold=false;										//deactivate the bold tag
			font.setBold(false);
		
		}else if (str.equalsIgnoreCase("[i]")){					//if line was [i]
			isItalic=true;										//activate the italic tag
																//and aplly it
			font.setItalic(true);
	
		}else if (str.equalsIgnoreCase("[/i]")){				//if line was [/i]
			isItalic=false;										//deactivate the italic tag
																//and aplly it
			font.setItalic(false);
		//if active, go to the tableEdit method for adding values to the current table
		}else if(TableCreation){
			
			tableEdit(str);
			
		//new page
		}else if(str.startsWith("[page]")){
			
			builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
		//comment that are not shown in the document	
		}else if(str.equalsIgnoreCase("[notshow]") || str.startsWith("//")){
			
			return;
		//print the unsupported tag	
		}else if(str.startsWith("[")){
			
			System.out.println("Tag not found: "+str);
		
		}else{													//if there was no tag
			
			builder.write(temp);								//add the text into current paragraph style
			
		}
		
	}

	/*
	 * Editing and adding values in the current table
	 */
	private void tableEdit(String str) throws Exception {
		
		//the next values will be headers
		if(str.equalsIgnoreCase("[headers]")){
			headers=true;
		//the next values stoped to be headers
		}else if (str.equalsIgnoreCase("[/headers]")){
			headers=false;
			builder.endRow();	//finish the current row of the table
		//the next values will be data
		}else if (str.equalsIgnoreCase("[data]")){
			data=true;
		//the next values stoped to be data
		}else if (str.equalsIgnoreCase("[/data]")){
			data=false;
		//number of columns of the table
		}else if(str.startsWith("[columns=")){
			columns=Integer.parseInt(str.substring(9, 10));
		//type of the table
		}else if(str.startsWith("[type=")){
			tableStyle=(str.substring(6,str.length()-1));
		//font size of table
		}else if(str.startsWith("[s=")){
			
		}else{
		
			// if the tag is header, all next values will 
			//be in the first line of table
			if(headers){
				builder.insertCell();	//new cell for each value in the same line
				builder.write(str);		//insert the value
			
			//if the tag is data, then will be in format :
			//	for example for 3 columns  data | data | data
			//  so we have to make tokens and insert each one in separate cells 
			// in the same line
			}else if (data){
				
				str=str.replaceAll(" ", "");	//clear white space
			     StringTokenizer st = new StringTokenizer(str,"|"); 	//make tokens
			     while (st.hasMoreTokens()) {	//insert all tokens
			         builder.insertCell();
			         builder.write(st.nextToken());
			         
			     }
			     builder.endRow();	//end the line
			}
			
			
		}
		
	}
	
	//this method search the style. See report for more info, or the documentation about setStyleIdentifier
	private void setTableStyle(String style){
		table= (Table)doc.getChild(NodeType.TABLE, tableID, true);
	
		try {
			table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);
		} catch (Exception e) {
			 
			e.printStackTrace();
		}
	
		if(style.equalsIgnoreCase("grid8")){
				try {
					table.setStyleIdentifier(133);
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}

			}else if(style.equalsIgnoreCase("classic2")){
				try {
			
					table.setStyleIdentifier(115);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("colorful2")){
				try {
					
					table.setStyleIdentifier(119);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("columns2")){
				try {
					
					table.setStyleIdentifier(122);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("columns3")){
				try {
					
					table.setStyleIdentifier(123);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("columns4")){
				try {
					
					table.setStyleIdentifier(124);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("columns5")){
				try {
					
					table.setStyleIdentifier(125);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("simple1")){
				try {
					
					table.setStyleIdentifier(112);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("simple2")){
				try {
					
					table.setStyleIdentifier(113);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("simple3")){
				try {
					
					table.setStyleIdentifier(114);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("subtle1")){
				try {
					
					table.setStyleIdentifier(148);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("subtle2")){
				try {
					
					table.setStyleIdentifier(149);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}else if(style.equalsIgnoreCase("classic")){
				try {
					
					table.setStyleIdentifier(114);    
				} catch (Exception e1) {
					 
					e1.printStackTrace();
				}
			}
	}
	
	//new paragraph
	public void newTag(int tag) throws Exception{
		par=new Paragraph(doc);
		paragraphFormat = builder.getParagraphFormat();
		font.setItalic(isItalic);
		paragraphFormat.setKeepTogether(true);
		font.setBold(isBold);
		paragraphFormat.setAlignment(tag);	//change the allignment
	
}
	//saves and creates the .doc with the given name 
	public void save(String filename){
	  
		try{
																				//opens the output stream to the file
			doc.save(filename+".docx");													//creates the file
																					//closing the stream
			
		}catch (Exception e){													//if somethings happends
			e.printStackTrace();												//print error
			
		}
	}

	
	
	
	
	
	
	//search for a color based on str
public static java.awt.Color findColor(String str){
	
		if (str.equalsIgnoreCase("green")){
			
			return java.awt.Color.GREEN;
		}else if(str.equalsIgnoreCase("blue")){
			return java.awt.Color.BLUE;
				
		}else if(str.equalsIgnoreCase("red")){
			
			return  java.awt.Color.RED;
			
		}else if(str.equalsIgnoreCase("black")){
			return java.awt.Color.BLACK;
				
		}else if(str.equalsIgnoreCase("gray")){
			return java.awt.Color.GRAY;
			
				
		}else if(str.equalsIgnoreCase("cyan")){
			return java.awt.Color.CYAN;
				
		}else if(str.equalsIgnoreCase("orange")){
			return  java.awt.Color.ORANGE;
				
		}else if(str.equalsIgnoreCase("pink")){
			return java.awt.Color.PINK;
				
			
		}else if(str.equalsIgnoreCase("white")){
			return java.awt.Color.WHITE;
			
		}else if(str.equalsIgnoreCase("green")){
			return java.awt.Color.GREEN;
			
		}else if(str.equalsIgnoreCase("yellow")){
			return java.awt.Color.YELLOW;
				
		}else if(str.equalsIgnoreCase("magenta")){
			return java.awt.Color.MAGENTA;
				
		}else if(str.equalsIgnoreCase("lightblue")){
			return java.awt.Color.BLUE;
				
		}else if(str.equalsIgnoreCase("darkgreen")){
			return java.awt.Color.darkGray;
		//if none in found, return white
		}else{
			return java.awt.Color.WHITE;	
		}
	}
}
