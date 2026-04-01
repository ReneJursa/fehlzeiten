package lib;

import javafx.scene.Node;
import javafx.scene.control.Label;
import javafx.scene.control.Tab;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.text.Font;
import javafx.scene.text.FontPosture;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;

import java.util.List;

public class AuthorTab extends Tab {
	TextField descField = new TextField();
	Text text = new Text();
	
	public AuthorTab(String s) {
		this.setText(s);
		init();
	}

	public void init() {
		
		String text_str = "\nRen\u00E9 Jursa\n" +
		              "\u25E6 Entwicklung\n" +
					  "\u25E6 Dokumentation\n" + 
					  "\u25E6 Kontakt: <renejursa2\u0040gmail.com>\n"; 
		text.setText(text_str);
		text.setFont(Font.font("Times New Roman", 13));
		GridPane gridPane = new GridPane();
		gridPane.addRow(0, text);
		this.setContent(gridPane);
	}
}
