import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;

import com.aspose.cells.Area;
import com.aspose.cells.Axis;
import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartArea;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartFrame;
import com.aspose.cells.ChartPoint;
import com.aspose.cells.ChartPointCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.Series;
import com.aspose.cells.SeriesCollection;
import com.aspose.cells.Style;
import com.aspose.cells.Title;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;



public class CChart {
	//Instantiating a Workbook object
	Workbook workbook ;
	//Obtaining the reference of the first worksheet
	WorksheetCollection worksheets;
	Worksheet sheet;
	boolean addValue;
	Cells cells;
	Cell cell;
	int i =1;
	int j =1;
	char column='A';
	char startingColumn='C';
	Chart chart;
	Title title;
	String temp;
	boolean addCategory;
	Color color;
	Area area;
	ArrayList<String> colorsArray;
	Series aSeries;
	int chartIndex ;
	ChartCollection charts;
	Style style;
	Font font;
	boolean yvalues ,xvalues;
	String type;
	
	public CChart(String type)  {
		 workbook = new Workbook();	//making a new Excel file
		worksheets = workbook.getWorksheets();
		sheet =  worksheets.get(0);	//get first sheet
		
		 charts = sheet.getCharts();
		 chartIndex = charts.add(ChartType.COLUMN,5,0,35,10);	//new default chart
		 chart = charts.get(chartIndex);	//get first chart index
		//Adding a chart to the worksheet
		colorsArray=new ArrayList<String>();		//the array for the colors of our bars/lines/etc
		 addCategory=false;
		addValue=false;
		yvalues=xvalues=false;
		this.type=type;		//the type of the chart
		
	}
	
	
	
	
	//main method for handling the chart
	public void edit(String str) throws  java.lang.StringIndexOutOfBoundsException{

		if(str.startsWith("[values") ){
			if(type.equals("[excel-scatter]")){
				i=1;		//starting with the first cell
				addValue=true;	//next inputs will be values
			}else{
				column++;	//start with column B
				i=1;		
				addValue=true;
			}
		}else if(str.startsWith("[/values")){
			addValue=false;	//next input will no longer be values
		
		//same logic for  values for x and y axis
		}else if(str.startsWith("[x-values")){
			column++;
			
			i=1;
			xvalues=true;
		}else if(str.startsWith("[y-values")){
			
			xvalues=false;
			column++;
			
			i=1;
		}else if(str.startsWith("[/x-values")){
			
			yvalues=true;
		}else if(str.startsWith("[/y-values")){
			yvalues=false;
		//if active, add the values into the right cell
		}else if(addValue){
			
				addValue(str);
		//category for the bar chart
		}else if(str.equalsIgnoreCase("[serieslabel]") ){
			
				addCategory=true;
		
		}else if(str.equalsIgnoreCase("[/serieslabel]")){
	
				addCategory=false;
		//adding the color into the array, so the first column will have the first value of array, etc
		}else if (str.startsWith("[column color=")){

			temp= str.substring(14,str.length()-1);
			colorsArray.add(temp);
			//adding the color into the array, so the first bar will have the first value of array, etc
		}else if(str.startsWith("[bar color=")){
			
			temp= str.substring(11,str.length()-1);
			colorsArray.add(temp);
			//adding the color into the array, so the first line will have the first value of array, etc
		}else if (str.startsWith("[line color=")){
			
			temp= str.substring(12,str.length()-1);
			colorsArray.add(temp);
			//if active, add the values into the right cell
		}else if(addCategory){
				
				addCategory(str);
				
		}else if(str.startsWith("[plot=")){
			ChartFrame plotArea = chart.getPlotArea();
			area = plotArea.getArea();
			temp= str.substring(6,str.length()-1);
			area.setForegroundColor(findColor(temp));
		//seting the title of the chart	
		}else if (str.startsWith("[title=")){
			
			temp= str.substring(7,str.length()-1);
			title = chart.getTitle();
			title.setText(temp);
			//seting the title of the X axis		
		}else if(str.startsWith("[xtitle=")){
			Axis categoryAxis = chart.getCategoryAxis();
			title = categoryAxis.getTitle();
			temp= str.substring(8,str.length()-1);
			title.setText(temp);
			//seting the title of the Y axis		
		}else if(str.startsWith("[ytitle=")){
			Axis valueAxis = chart.getValueAxis();
			title = valueAxis.getTitle();
			temp= str.substring(8,str.length()-1);
			title.setText(temp);
			
			
		}else if (str.startsWith("[plot by ")){
		
		//seting the color of outside of the chart
		}else if(str.startsWith("[chart color=")){
			
			ChartArea chartArea = chart.getChartArea();
			temp= str.substring(13,str.length()-1);
			area = chartArea.getArea();
			area.setForegroundColor(findColor(temp));
		//setting the color of the inside of the chart
		}else if(str.startsWith("[area=")){
			
			temp= str.substring(6,str.length()-1);
			
			if(type.equals("[excel-chart]")){
				ChartArea chartArea = chart.getChartArea();
				area = chartArea.getArea();
				area.setForegroundColor(findColor(temp));
			}else{
				ChartFrame plotArea = chart.getPlotArea();
				Area area = plotArea.getArea();
				area.setForegroundColor(findColor(temp));
			}
		//setting the chart type with the method setType
		}else if (str.startsWith("[chart type=")){
			temp = str.substring(12,str.length()-1);
			chart.setType(setChartType(temp));
		
		}else if(str.startsWith("[border=")){
		
		}else if (str.equalsIgnoreCase("[y-values]")){
			yvalues=true;
		//comments that will not show
		}else if(str.startsWith("//")){
			return;
		
		//unknown tag
		}else{
		
			System.out.println("Unknown tag: "+str);
		}
	}
	
	
	//search for the chart type based on str
	private int setChartType(String str){
		
		
		if (str.startsWith("column") || str.startsWith("xlColumnClustered")){

			return ChartType.COLUMN;
		}else if(str.equalsIgnoreCase("cylinder")){
			return ChartType.CYLINDER;
		
		}else if (str.equalsIgnoreCase("smoothNoMarkers") || str.equalsIgnoreCase("xlLineMarkers")){
		
			return ChartType.SCATTER_CONNECTED_BY_CURVES_WITHOUT_DATA_MARKER;
			
		
		
		}else{
			System.out.println("Unknown chart type: "+str);
			return 0;
		}
	}
	
	
	
	//insert the values of the chart for each category
	private void addValue(String str) {
		cells = sheet.getCells();
		cell = cells.get(column+""+i);	//for example the first value of first category will be B1, then B2 etc
	
		
		str=str.replaceAll(" ", "");	//clear white space
	
			//parse integer or double and insert it into the cell
		        try {
		        	
			        	cell.setValue(Integer.parseInt(str));
		            
		        } catch (NumberFormatException e) {
		       
		        	cell.setValue(Double.parseDouble(str));
		        }
		    
	
		i++;	//go to next cell
	}
	
	//insert the categories of the chart
	private void addCategory(String str) {
		cells = sheet.getCells();
		cell = cells.get("A"+j);	//for example, if there are 3 categories, they will go A1, A2, A3
		cell.setValue(str);		//insert the category into the cell
		j++;	//number of next catogory
		
	}
	
	SeriesCollection serieses;
	SeriesCollection nSeries;
	SeriesCollection SecondSerieses;
	SeriesCollection SecondNSeries;
	Axis axis;
	
	public void save(){
		style = cell.getStyle();
		font = style.getFont();
		font.setBold(DocConverter.isBold);
		
		 serieses = chart.getNSeries();	
		 //if scatter chart
		 if((type.equals("[excel-scatter]"))){
			
			 nSeries = chart.getNSeries();
			 serieses.add("C1:C19", true);
			chart.getNSeries().setCategoryData("B1:B19");
		//if not scatter, the values will be from B1 to the last column
		 }else{
		 	
			 serieses.add("B1:"+column+(i-1), true);
			 nSeries = chart.getNSeries();
		}
		 
		if(!(type.equals("[excel-scatter]"))){
			
			nSeries.setCategoryData("A1:A"+(j-1));
		}
		
		
		
		//adding color for each bar/line/etc
		for (int k=0;k<colorsArray.size();k++){
			Series aSeries = nSeries.get(0);
			area = aSeries.getArea();
			area.setForegroundColor(findColor(colorsArray.get(k)));
			
			//Setting the foreground color of the area of the 1st NSeries point
			ChartPointCollection chartPoints =aSeries.getPoints();
			ChartPoint point = chartPoints.get(k);
			point.getArea().setForegroundColor(findColor(colorsArray.get(k)));
			
		}	
		
		style = cell.getStyle();
		font = style.getFont();
		font.setBold(DocConverter.isBold);
		
		cell.setStyle(style);	//setting the style
		
		chart.setShowDataTable(false);	//hide data table
		//make the default option for image
		ImageOrPrintOptions imgOpts = new ImageOrPrintOptions();
		imgOpts.setImageFormat(ImageFormat.getPng());
	
		//Save the chart image file.
		try {
			chart.toImage(new FileOutputStream("temp.png"), imgOpts);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	
		
	}
	
	//find and return the Aspose color for the charts
public Color findColor(String str){
		
		if (str.equalsIgnoreCase("green")){
			
			return Color.getGreen();
		}else if(str.equalsIgnoreCase("blue")){
			return Color.getBlue();
				
		}else if(str.equalsIgnoreCase("red")){
			return  Color.getRed();
			
		}else if(str.equalsIgnoreCase("black")){
			return Color.getBlack();
				
		}else if(str.equalsIgnoreCase("gray")){
			return Color.getGray();
				
		}else if(str.equalsIgnoreCase("brown")){
			return Color.getBrown();
				
		}else if(str.equalsIgnoreCase("cyan")){
			return Color.getCyan();
				
		}else if(str.equalsIgnoreCase("orange")){
			return Color.getOrange();
				
		}else if(str.equalsIgnoreCase("pink")){
			return Color.getPink();
				
		}else if(str.equalsIgnoreCase("violet")){
			return Color.getViolet();
			
		}else if(str.equalsIgnoreCase("white")){
			
			return Color.getWhite();
			
		}else if(str.equalsIgnoreCase("green")){
			return Color.getGreen();
			
		}else if(str.equalsIgnoreCase("yellow")){
			return Color.getYellow();
				
		}else if(str.equalsIgnoreCase("magenta")){
			return Color.getMagenta();
				
		}else if(str.equalsIgnoreCase("lightblue")){
			return Color.getSkyBlue();
				
		}else if(str.equalsIgnoreCase("darkgreen")){
			return Color.getDarkGreen();
			
		}else if(str.equalsIgnoreCase("turquoise")){
			return Color.getTurquoise();
		}else{
			return Color.getWhite();	
		}
	}
}
