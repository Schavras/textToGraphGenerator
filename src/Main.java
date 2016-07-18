import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;


public class Main {

	public static void main(String[] args) throws Exception {
	
		File file = null;						//the opening file
		BufferedReader reader = null;			//the buffer to open the text
		
		
		String line="";							//temp line to read in every loop
		DocConverter docC = new DocConverter(); //the converter that handles the .docx
		
		
		try{
			file = new File("text.txt");		//you have to enter manually the path to the text
		
		} catch (NullPointerException e){
			System.err.println("File not found!!");									//if cannot be found, show error
		}
		
		try {
			reader = new BufferedReader (new InputStreamReader (new FileInputStream(file)));	//open the stream to read the text
		} catch (FileNotFoundException e ) {
			System.err.println("Error opening file!");											//if cannot open, show error
		}
		
		try {	
			System.out.println("Converting...");
			line = reader.readLine();				//try to read the first line
			while(line!=null){						//while there are more lines, continue
				docC.search(line);					//import every line into .doc, through this method
					
				line = reader.readLine();			//read next line
			}
		}catch (IOException e) {
			System.out.println("Line input error: Sudden end.");	// Error message in the reading of the line.
		}
	
	
		
	docC.save("Document");													//in the end, make and save the .doc file
	/*try {
	    Files.delete(Paths.get("temp.png"));
	} catch (NoSuchFileException x) {
	    System.err.format("temp.png");
	} catch (DirectoryNotEmptyException x) {
	    System.err.format("temp.png");
	} catch (IOException x) {
	    // File permission problems are caught here.
	    System.err.println(x);
	}*/
	System.out.println("Completed.");		//successful finished.
	}

}
