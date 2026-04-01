package lib;

import javafx.scene.Node;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Tab;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.text.Font;
import javafx.scene.text.FontPosture;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.net.URISyntaxException;

public class LicenseTab extends Tab {
	TextField descField = new TextField();
	Text text = new Text();
	
	public LicenseTab(String s) {
		this.setText(s);
		init();
	}

	public void init() {
		String text_str = "";
		try {
			//File fileCurDir = new File(getClass().getProtectionDomain().getCodeSource().getLocation().getPath());
			URL url = getClass().getProtectionDomain().getCodeSource().getLocation();
			File fileCurDir;
			try {
				fileCurDir = new File(url.toURI()); 
			} catch (URISyntaxException e) {
				fileCurDir = new File(url.getPath()); // Fallback
			}
			String curDir = fileCurDir.getName();
			Path parentPath = Paths.get(fileCurDir.getParent().toString());
			String parentPath_str = parentPath.toFile().toString();
			
			String fileStr = "LICENSE.txt";
			String pathToFile = "";
			
			if (Files.exists(Paths.get(fileStr))) {
				//System.out.println(fileStr + " in cur dir");
				pathToFile = parentPath_str + "/" + fileStr;
				//System.out.println(pathToFile);
			}
			else {
				Path parentParentPath = parentPath.getParent();
				String parentParentPath_str = parentParentPath.toFile().toString();
				pathToFile = parentParentPath_str + "/" + fileStr;
				/*
				if (Files.exists(Paths.get(pathToFile))) {
					System.out.println(fileStr + " in parent dir");
					System.out.println(pathToFile);
				}
				else {
					System.out.println(fileStr + " also not in parent dir");
				}
				*/
			}
			InputStreamReader in = new InputStreamReader(new FileInputStream(pathToFile), StandardCharsets.UTF_8);
			int c;
			StringBuilder sb = new StringBuilder();
			while ((c = in.read()) != -1) {
				sb.append((char) c);
			}
			text_str = sb.toString();
			//System.out.println(text_str);
		} catch (FileNotFoundException ex) {
			System.out.println(ex);
		} catch (IOException ex) {
			System.out.println(ex);
		}			
		
		text.setText(text_str);
		text.setFont(Font.font("Times New Roman", 13));
		GridPane gridPane = new GridPane();
		
		gridPane.addRow(0, text);
		ScrollPane sp = new ScrollPane(gridPane);
		this.setContent(gridPane);
	}
}
