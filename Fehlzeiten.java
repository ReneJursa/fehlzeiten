import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.net.URISyntaxException;
import java.util.Scanner;
import java.io.FileWriter;
import java.io.PrintWriter;
import java.awt.Color;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Arrays;
import java.util.Map;
import java.util.TreeMap;
import java.util.Comparator;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.SortedList;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.geometry.VPos;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.TextField;
import javafx.scene.control.TabPane;
import static javafx.scene.control.TabPane.TabClosingPolicy;
import javafx.scene.control.ScrollPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
import javafx.stage.Stage;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import static javafx.stage.FileChooser.ExtensionFilter;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.text.Text;
import javafx.scene.text.FontWeight;

import org.openpdf.text.Document;
import org.openpdf.text.DocumentException;
import org.openpdf.text.Element;
import org.openpdf.text.PageSize;
import org.openpdf.text.Font;
import org.openpdf.text.Chunk;
import org.openpdf.text.Anchor;
import org.openpdf.text.List;
import org.openpdf.text.ListItem;
import org.openpdf.text.Paragraph;
import org.openpdf.text.HeaderFooter;
import org.openpdf.text.Rectangle;
import org.openpdf.text.Phrase;
import org.openpdf.text.pdf.PdfPCell;
import org.openpdf.text.pdf.PdfPTable;
import org.openpdf.text.pdf.PdfWriter;
import org.openpdf.text.pdf.PdfPCellEvent;
import org.openpdf.text.pdf.PdfContentByte;
import org.openpdf.text.pdf.draw.LineSeparator;

import lib.Abwesenheit;
import lib.AuthorTab;
import lib.DescriptionTab;
import lib.DocumentationTab;
import lib.LicenseTab;
import lib.PercentCell;

public class Fehlzeiten extends Application {
	
	private boolean tM_File_imported = false;
	private boolean excuseEdited = false;
	private boolean aps_File_imported = false;
	private boolean aps_File_mustBeRead = false;
	private boolean kursUndAnfangsDatumFlag = false;
	private boolean initialDirSelected = false;
	
	private TreeMap<String, Abwesenheit> map = new TreeMap<String, Abwesenheit>();
		
	private TreeMap<String, TreeMap<String, Abwesenheit>> map2 = new TreeMap<String, TreeMap<String, Abwesenheit>>();
		
	private TreeMap<String, TreeMap<String, Abwesenheit>> map3 = new TreeMap<String, TreeMap<String, Abwesenheit>>();
		
	private TreeMap<String, Integer> mapNrOfLessonsPerSubject = new TreeMap<String, Integer>();
	
	private ArrayList<String> initExcuseList = new ArrayList<String>();
	
	private ObservableList<String> obsExcuseList;
	
	private ListView<String> lvExcuseList;
	
	private int nrOfLessons = 0;
	
	private String startDate = "";
	private String endDate = "";
	private String kurs = "";
	
	private String tM_str = "";
	private String aps_str = ""; 
	
	private String styleAnchorPaneStr = "-fx-background-color: snow;";
	
	private String initialDir = "";
	
	private String resourcesDir_str  = "";
	
	private final String fzVersionStr = "FEHLZEITEN 1.4.0";
	
	private Stage mainStage;
	private Stage lastOpenMainStage = null;
	
	private Stage excuseStage;
	private Stage lastOpenExcuseStage = null;
	
	private Stage infoStage;
	
	private Alert alert = new Alert(AlertType.NONE);
	
	public void showErrorAlert(String str) {
		alert.setAlertType(AlertType.ERROR);
		alert.setContentText(str);
		alert.showAndWait();
	}
	public void showInformationAlert(String str) {
		alert.setAlertType(AlertType.INFORMATION);
		alert.setContentText(str);
		alert.showAndWait();
	}
	public void showWarningAlert(String str) {
		alert.setAlertType(AlertType.WARNING);
		alert.setContentText(str);
		alert.showAndWait();
	}
	
	public static String replaceToUniCode(String str) {
		str = str.replace("Ä","\u00C4");
		str = str.replace("ä","\u00E4");
		str = str.replace("Ö","\u00D6");
		str = str.replace("ö","\u00F6");
		str = str.replace("Ü","\u00DC");
		str = str.replace("ü","\u00FC");
		str = str.replace("ß","\u00DF");
		str = str.replace("à","\u00E0");
		str = str.replace("á","\u00E1");
		str = str.replace("è","\u00E8");
		str = str.replace("é","\u00E9");
		str = str.replace("ì","\u00EC");
		str = str.replace("í","\u00ED");
		str = str.replace("ò","\u00F2");
		str = str.replace("ó","\u00F3");
		str = str.replace("ù","\u00F9");
		str = str.replace("ú","\u00FA");
		return str;
	}
	
	private double round(double value, int decimalPoints) {
		double d = Math.pow(10, decimalPoints);
		return Math.round(value * d) / d;
    }
	
	private boolean obsExcuseListContainsIgnoreCase(String soughtAfter) {
		for (String current: obsExcuseList) {
			if (current.equalsIgnoreCase(soughtAfter)) {
				return true;
			}
		}
		return false;
	}
	
	@Override
	public void start(Stage stage) {
		mainStage = new Stage();
		excuseStage = new Stage();
		infoStage = new Stage();
		createMainStage(stage);
	}
	
	public static void main(String[] args) {
		launch(args);
	}
	
	private void ready() {
		Platform.exit();  
	}
	
	private void about() {
		
		Text fzVersionText = new Text(fzVersionStr);
		fzVersionText.setFont(javafx.scene.text.Font.font("Helvetica", FontWeight.BOLD, 24));
		
		fzVersionText.setTextOrigin(VPos.TOP);
		DescriptionTab descTab = new DescriptionTab("Info");
		AuthorTab authorTab = new AuthorTab("Autor");
		LicenseTab licenseTab = new LicenseTab("Lizenz");
		DocumentationTab documentationTab = new DocumentationTab("Dokumentation");

		TabPane tabPane = new TabPane();
		tabPane.setTabClosingPolicy(TabClosingPolicy.UNAVAILABLE);

		tabPane.getTabs().addAll(descTab, authorTab, licenseTab, documentationTab);
		
		HBox hBox = new HBox(tabPane);	
		hBox.setStyle("-fx-padding: 10;");
		
		GridPane gridPane = new GridPane();
		gridPane.setVgap(10);
		gridPane.setHgap(10);
		
		gridPane.add(fzVersionText, 0, 0);
		
		gridPane.add(hBox, 0, 1);
		
		AnchorPane anchorPane = new AnchorPane();
		anchorPane.setStyle(styleAnchorPaneStr);
		anchorPane.getChildren().add(gridPane);
		anchorPane.setPrefWidth(770);
		anchorPane.setPrefHeight(650);
		
		AnchorPane.setLeftAnchor(gridPane, 30.);
		AnchorPane.setTopAnchor(gridPane, 30.);
		ScrollPane sp = new ScrollPane(anchorPane);
		//System.out.println("Vvalue: " + sp.getVvalue());
		sp.setVvalue(-50.0);
		Scene scene = new Scene(sp, 770, 650);
		
		infoStage.setScene(scene);
		infoStage.setTitle("\u00DCber - FEHLZEITEN");
		infoStage.show();
		infoStage.toFront();
	}
	
	private void createMainStage(Stage stage) {
		
		String statEntStr = "entsch.";
		String aubStr = "AUB";
		String attStr = "Attest";
		String krankStr = "Krankheit";
		String entStr = "Entschuldigt";
		String behStr = "Beh\u00F6rde";
		
		initExcuseList.add(statEntStr);
		initExcuseList.add(aubStr);
		initExcuseList.add(attStr);
		initExcuseList.add(krankStr);
		initExcuseList.add(entStr);
		initExcuseList.add(behStr);
		
		obsExcuseList = FXCollections.observableArrayList(initExcuseList);
		
		System.out.println("voreingestellte Entschuldigungsgr\u00FCnde:");
		for(String e: obsExcuseList) {
			System.out.println(e);
		}
		
		AnchorPane anchorPane = new AnchorPane();
		GridPane pane = new GridPane();
		pane.setHgap(50);
		pane.setVgap(10);
		
		Label infoLabel = new Label("\u00DCber FEHLZEITEN");
		infoLabel.setStyle("-fx-font-family: Helvetica;" +  
						   "-fx-font-size: 16px;" +
						   "-fx-font-weight: bold;" +
						   "-fx-text-fill: black");
		
		pane.add(infoLabel, 0, 0);
		
		Button infoBtn = new Button();
		infoBtn.setText("\u00DCber");
		infoBtn.setOnAction(event -> about());
		pane.add(infoBtn, 0, 1);
		
		Label excuseLabel = new Label("\nBearbeiten der akzeptierten\nEntschuldigungsgr\u00FCnde");
		excuseLabel.setStyle("-fx-font-family: Helvetica;" +  
						   "-fx-font-size: 13px;" +
						   "-fx-font-weight: bold;" +
						   "-fx-text-fill: black");
		
		pane.add(excuseLabel, 0, 2);
		
		Button excuseBtn = new Button();
		excuseBtn.setText("Bearbeiten");
		excuseBtn.setOnAction(event -> editExcuseList());
		pane.add(excuseBtn, 0, 3);
		
		Label dirLabel = new Label("\n\nOrdner ausw\u00E4hlen");
		dirLabel.setStyle("-fx-font-family: Helvetica;" +  
						   "-fx-font-size: 13px;" +
						   "-fx-font-weight: bold;" +
						   "-fx-text-fill: black");
		
		pane.add(dirLabel, 0, 4);
		
		Button dirBtn = new Button();
		dirBtn.setText("W\u00E4hle Ordner");
		dirBtn.setOnAction(event -> selectDir(stage));
		pane.add(dirBtn, 0, 5);
		
		Label loadLabel = new Label("\n\u00D6ffnen von TeachingMethods als CSV-Datei");
		loadLabel.setStyle("-fx-font-family: Helvetica;" +  
						   "-fx-font-size: 13px;" +
						   "-fx-font-weight: bold;" +
						   "-fx-text-fill: black");
		
		pane.add(loadLabel, 0, 6);
		
		Button loadBtn = new Button();
		loadBtn.setText("Datei \u00F6ffnen");
		loadBtn.setOnAction(event -> readTeachingMethodsFile(stage));
		pane.add(loadBtn, 0, 7);
		
		Label loadLabel2 = new Label("\n\u00D6ffnen von AbsencePerStudent als CSV-Datei");
		loadLabel2.setStyle("-fx-font-family: Helvetica;" +  
						   "-fx-font-size: 13px;" +
						   "-fx-font-weight: bold;" +
						   "-fx-text-fill: black");
		
		pane.add(loadLabel2, 0, 8);
		
		Button loadBtn2 = new Button();
		loadBtn2.setText("Datei \u00F6ffnen");
		loadBtn2.setOnAction(event -> readAbsencePerStudentFile(stage));
		pane.add(loadBtn2, 0,9);
		
		Label calcLabel = new Label("\nErstellen der Fehlzeitentabellen\nals CSV-Datei");
		calcLabel.setStyle("-fx-font-family: Helvetica;" +  
						   "-fx-font-size: 13px;" +
						   "-fx-font-weight: bold;" +
						   "-fx-text-fill: black");
		
		pane.add(calcLabel, 1, 2);
		
		Button calcBtn = new Button();
		calcBtn.setText("Erstellen");
		calcBtn.setOnAction(event -> erstelleFehlzeitenTabelle(stage));
		pane.add(calcBtn, 1, 3);
		
		Label pdfLabel = new Label("\nErstellen der Fehlzeitentabellen\nals PDF-Datei");
		pdfLabel.setStyle("-fx-font-family: Helvetica;" +  
						   "-fx-font-size: 13px;" +
						   "-fx-font-weight: bold;" +
						   "-fx-text-fill: black");
		
		pane.add(pdfLabel, 1, 4);
		
		Button pdfBtn = new Button();
		pdfBtn.setText("Erstellen");
		pdfBtn.setOnAction(event -> createPDF(stage));
		pane.add(pdfBtn, 1, 5);
		
		Label endLabel = new Label("\nProgramm beenden");
		endLabel.setStyle("-fx-font-family: Helvetica;" +  
					   "-fx-font-size: 13px;" +
					   "-fx-font-weight: bold;" +
					   "-fx-text-fill: black");
		pane.add(endLabel, 1, 8);
		
		Button endBtn = new Button();
		endBtn.setText("Fertig");
		endBtn.setOnAction(event -> ready());
		pane.add(endBtn, 1, 9);
		
		anchorPane.setStyle(styleAnchorPaneStr);
		
		anchorPane.getChildren().add(pane);
		anchorPane.setPrefWidth(610);
		anchorPane.setPrefHeight(520);
		AnchorPane.setLeftAnchor(pane, 30.);
		AnchorPane.setTopAnchor(pane, 30.);
		
		Scene scene = new Scene(anchorPane, 610, 520);
		mainStage.setScene(scene);
		mainStage.setTitle("FEHLZEITEN");
		
		if(this.lastOpenMainStage == null) {
			mainStage.setX(200);
			mainStage.setY(100);
		}
		else {
			mainStage.setX(this.lastOpenMainStage.getX());
			mainStage.setY(this.lastOpenMainStage.getY());
		}
		mainStage.setMinWidth(610);
		mainStage.setMinHeight(520);
		mainStage.show();
		this.lastOpenMainStage = mainStage;
	}
	
	private void editExcuseList() {
	    
		aps_File_mustBeRead = true;
		excuseEdited = false;
		
		obsExcuseList = FXCollections.observableArrayList(initExcuseList);
        
		SortedList<String> excuseListSorted = new SortedList<>(obsExcuseList, Comparator.naturalOrder());
        lvExcuseList = new ListView<>(excuseListSorted);

		Label lblReason = new Label("Grund:");
		Label lblReasons = new Label("Gr\u00FCnde:");
		TextField tfReason = new TextField();
		Button btnAdd = new Button("Einf\u00FCgen");
		Button btnRemove = new Button("L\u00F6schen");

        btnAdd.setOnAction(event -> {
            String s = replaceToUniCode(tfReason.getText());
            if (s.length() > 0  &&  !obsExcuseList.contains(s))
                obsExcuseList.add(s);
				tfReason.setText(null);
				aps_File_mustBeRead = true;
				excuseEdited = true;
        });

        btnRemove.setOnAction(event -> {
                obsExcuseList.remove(lvExcuseList.getSelectionModel().getSelectedItem());
				aps_File_mustBeRead = true;
				excuseEdited = true;
		});

        GridPane root = new GridPane();

        double dist = 10.0;
        root.setPadding(new Insets(dist, dist, dist, dist));
        root.setHgap(dist);
        root.setVgap(dist);

        root.add(lblReason, 0, 0); root.add(tfReason, 1, 0); root.add(btnAdd, 2, 0);
        root.add(lblReasons, 0, 1); root.add(lvExcuseList, 1, 1); root.add(btnRemove, 2, 1);

        GridPane.setHalignment(btnAdd, HPos.RIGHT);
        GridPane.setHalignment(btnRemove, HPos.RIGHT);

        GridPane.setHgrow(tfReason, Priority.ALWAYS);
        GridPane.setVgrow(lvExcuseList, Priority.ALWAYS);

        excuseStage.setTitle("Entschuldigungsgr\u00FCnde - FEHLZEITEN");
        excuseStage.setScene(new Scene(root, 450, 300));
		if(this.lastOpenExcuseStage == null) {
			excuseStage.setX(450);
			excuseStage.setY(400);
		}
		else {
			excuseStage.setX(this.lastOpenExcuseStage.getX());
			excuseStage.setY(this.lastOpenExcuseStage.getY());
		}
        excuseStage.setMinWidth(400);
        excuseStage.setMinHeight(300);
        excuseStage.show();
		this.lastOpenExcuseStage = excuseStage;
    }
	
	private void selectDir(Stage stage) {
		try {
			DirectoryChooser directoryChooser = new DirectoryChooser();
       
			directoryChooser.setTitle("W\u00E4hle einen Ordner aus");
		
			File selectedDirectory = directoryChooser.showDialog(stage);
		
			initialDir = selectedDirectory.getAbsolutePath();
			System.out.println("Ordner: " + initialDir);
			initialDirSelected = true;
		} catch (NullPointerException ex) {
			System.err.println(ex); 
		}
	}
	
	private void readTeachingMethodsFile(Stage stage) {
		if (!initialDirSelected) {
			showInformationAlert("Es wurde noch kein Ordner ausgew\u00E4hlt");  
			return;
		}
		mapNrOfLessonsPerSubject.clear();
		
		kursUndAnfangsDatumFlag = false;
		
		FileChooser fileChooser = new FileChooser();
		fileChooser.setTitle("Lese TeachingMethods-Datei");
		fileChooser.getExtensionFilters().add(new ExtensionFilter("CSV", "*.csv"));
		fileChooser.setInitialDirectory(new File(initialDir));
		File tM_File = fileChooser.showOpenDialog(stage);
		if (tM_File == null) {
			showWarningAlert("Keine Datei ausgew\u00E4hlt oder Datei nicht lesbar");  
			return;
		}
		tM_str = tM_File.getName();
		System.out.println("\nTeachingMethods-Datei: " + tM_str);
		
		try {
			Scanner inp_tM = new Scanner(tM_File,"UTF-8");
			String storno = "Storniert";
			nrOfLessons = 0;
			String subject = "";
			String inp_header = replaceToUniCode(inp_tM.nextLine());
			//System.out.println(inp_header);
			int value = 0;
			int i = 1;
			while(inp_tM.hasNext()) { 			
				String inp_zeile = inp_tM.nextLine();
				String zeile = inp_zeile.replace('\u0022', '\u0020');
				zeile = zeile.replace('\t', ';');
				//System.out.println(i + " " + zeile);
				String s[]   = zeile.trim().split(";");
				if (s.length < 3) {
					i++;
					continue;
				}
				//System.out.println(i + " " + zeile);
				if (s[2] != null && !s[2].equals("")) {
					endDate = s[2];
				}
				boolean lesson = true;
				if (s.length > 6 && s[6] != null && storno.equalsIgnoreCase(s[6])) {
					lesson = false;
				}
				if (lesson) {
					if (s.length > 3 && s[3] != null) {
						value = Integer.parseInt(s[3]);
					}
					else {
						value = 0;
					}
					nrOfLessons += value;
					if (s.length > 4) {
						if (s[4] != null && !s[4].equals("")) {
							subject = replaceToUniCode(s[4]);
						}
						else {
							subject = "unbekannt";
							//System.out.println("kein Fach");
							//System.out.println(i + " " + zeile);
						}
					}
					else {
						subject = "unbekannt";
						//System.out.println("kein Fach 2");
						//System.out.println(i + " " + zeile);
					}
					if (mapNrOfLessonsPerSubject.containsKey(subject)) {
						int value_prev = mapNrOfLessonsPerSubject.get(subject);
						int value_new =  value_prev + value;
						mapNrOfLessonsPerSubject.put(subject, value_new);
					}
					else {
						if (value > 0) {
							mapNrOfLessonsPerSubject.put(subject, value);
						}
					}		
				}
				if (s[0] != null && !s[0].equals("") && s[2] != null && !s[2].equals("") && !kursUndAnfangsDatumFlag) {
					kurs = s[0];
					startDate = s[2];
					kursUndAnfangsDatumFlag = true;
				}
				i++;
			}
			inp_tM.close();
			//System.out.println("\nAnzahl eingelesener Zeilen der TeachingMethods-Datei: " + i);
			if (kursUndAnfangsDatumFlag == false) {
				System.out.println("Die Klasse und das Anfangsdatum der TeachingMethods-Datei " + tM_str + "\nkonnten nicht bestimmt werden");
				showErrorAlert("Die Klasse und das Anfangsdatum der\nTeachingMethods-Datei " + tM_str + "\nkonnten nicht bestimmt werden");
				mapNrOfLessonsPerSubject.clear();
				return;
			}
			System.out.println("Klasse: " + kurs);
			System.out.println("Anfangsdatum: " + startDate + " Enddatum: " + endDate + " des ausgew\u00E4hlten Zeitbereichs"); 
			System.out.println("Gesamtzahl der Unterrichtsstunden im ausgew\u00E4hlten Zeitbereich: " + nrOfLessons);
			
			for (Map.Entry<String, Integer> mapIt : mapNrOfLessonsPerSubject.entrySet()) {
				System.out.printf("Fach: %s U-Std: %d\n", mapIt.getKey(), mapIt.getValue());
			} 
			showInformationAlert(tM_str + "\nwurde eingelesen.");
			tM_File_imported = true;
			aps_File_mustBeRead = true;
		
		} catch (FileNotFoundException ex) {
			System.err.println(ex);
			showErrorAlert("Die TeachingMethods-Datei " + tM_str + " konnte nicht eingelesen werden\n" + ex);
			mapNrOfLessonsPerSubject.clear();
		} catch (NumberFormatException ex) {
			System.err.println(ex);
			showErrorAlert("Die TeachingMethods-Datei " + tM_str + " konnte nicht eingelesen werden\n" + ex);
			mapNrOfLessonsPerSubject.clear();
		}	
	}
	
	private void readAbsencePerStudentFile(Stage stage) {
		
		map.clear();
		map2.clear();
		map3.clear();
		
		if(tM_File_imported == false) {
			showInformationAlert("Es wurde noch keine TeachingMethods-Datei eingelesen.");
			return;
		}
		if(mapNrOfLessonsPerSubject.isEmpty()) {
			showInformationAlert("Die TeachingMethods-Datei muss eingelesen werden.");
			return;
		}
		if(excuseEdited == false) {
			showInformationAlert("Voreingestellte Entschuldigungsgr\u00FCnde werden verwendet.");
		}
		FileChooser fileChooser = new FileChooser();
		fileChooser.setTitle("Lese AbsencePerStudent-Datei");
		fileChooser.getExtensionFilters().add(new ExtensionFilter("CSV", "*.csv"));
		fileChooser.setInitialDirectory(new File(initialDir));
		File aps_File = fileChooser.showOpenDialog(stage);
		if (aps_File == null) {
			showWarningAlert("Keine Datei ausgew\u00E4hlt oder Datei nicht lesbar");  
			return;
		}
		aps_str = aps_File.getName();
		System.out.println("\nAbsencePerStudent-Datei: " + aps_str);
		
		try {
			Scanner inp_aps = new Scanner(aps_File,"UTF-8");
			String name = "";
			String kursAbs = "";
			String subject = "";
			String status = "";
			String reason = "";
			String reason2 = "";
			String reason3 = "";
			
			System.out.println("\nverwendete Entschuldigungsgr\u00FCnde:");
			for(String e: obsExcuseList) {
				System.out.println(e);
			}
			String inp_header = replaceToUniCode(inp_aps.nextLine());
			//System.out.println(inp_header);
			int value = 0;
			int i = 1;
			while(inp_aps.hasNext()) { 			
				String inp_zeile = inp_aps.nextLine();
				String zeile = inp_zeile.replace('\u0022', '\u0020');
				zeile = zeile.replace('\t', ';');
				//System.out.println(i + " " + zeile);
				String s[]   = zeile.trim().split(";");
				if (s.length <= 17) {
					System.out.println("Keine g\u00FCltige Zeile in Zeile " + (i+1) + ":");
					System.out.println(zeile);
					//showErrorAlert(i+1 + " Keine g\u00FCltige Zeile");
					i++;
					continue;
				}
				if (s[0] != null && !s[0].equals("")) {
					name = replaceToUniCode(s[0]);
				}
				else {
					showWarningAlert("Kein Name in Zeile " + (i+1)); 
					name = "unbekannt";
					System.out.println("kein Name in AbsencePerStudent-Datei in Zeile " + (i+1) + ":");
					System.out.println(zeile);
				}
				if (s[3] != null && !s[3].equals("")) {
					kursAbs = replaceToUniCode(s[3]);
				}
				else {
					kursAbs = "unbekannt";
				}
				if (!kursAbs.equals(kurs)) {
					showWarningAlert("Andere Klasse " + kursAbs + " in Zeile " + (i+1) + " bei " + name);
					System.out.println("Andere Klasse " + kursAbs + " in AbsencePerStudent-Datei in Zeile " + (i+1) + ":");
					System.out.println(zeile);
					i++;
					continue;
				}
				if (s[7] != null) {
					value = Integer.parseInt(s[7]);
				}
				else {
					value = 0;
				}
				if (s[9] != null && !s[9].equals("")) {
					subject = replaceToUniCode(s[9]);
				}
				else {
					showWarningAlert("Kein Fach in Zeile " + (i+1) + " bei " + name); 
					subject = "unbekannt";
					System.out.println("kein Fach in AbsencePerStudent-Datei in Zeile " + (i+1) + ":");
					System.out.println(zeile);
				}
				if(!mapNrOfLessonsPerSubject.containsKey(subject)) {
					showWarningAlert("Anderes Fach " + subject + " in Zeile " + (i+1) + " bei " + name);
					System.out.println("Anderes Fach " + subject + " in AbsencePerStudent-Datei in Zeile " + (i+1) + ":");
					System.out.println(zeile);
					i++;
					continue;
				}
				if (s[10] != null) {
					reason = replaceToUniCode(s[10]);
				}
				if (s[11] != null) {
					reason2 = replaceToUniCode(s[11]);
				}
				if (s[15] != null) {
					reason3 = replaceToUniCode(s[15]);
				}
				if (s[17] != null) {
					status = replaceToUniCode(s[17]);
				}
				
				boolean entschuldigt = false;
				
				if (obsExcuseListContainsIgnoreCase(status) ||
					obsExcuseListContainsIgnoreCase(reason) ||
					obsExcuseListContainsIgnoreCase(reason2) ||
					obsExcuseListContainsIgnoreCase(reason3)) {
						
					entschuldigt = true; 
				}
				if (map.containsKey(name)) { 
					if (entschuldigt) {
						int value_prev = map.get(name).getEntAbw();
						int value_new = value_prev + value;
						map.get(name).setEntAbw(value_new);
					}
					else {
						int value_prev = map.get(name).getUnEntAbw();
						int value_new = value_prev + value;
						map.get(name).setUnEntAbw(value_new);
					}	
					if (map2.get(name).containsKey(subject)) {
						if (entschuldigt) {
							int value_prev = map2.get(name).get(subject).getEntAbw();
							int value_new =  value_prev + value;
							map2.get(name).get(subject).setEntAbw(value_new);
						}
						else {
							int value_prev = map2.get(name).get(subject).getUnEntAbw();
							int value_new =  value_prev + value;
							map2.get(name).get(subject).setUnEntAbw(value_new);
						}
						//System.out.printf("Fach %s vp %d  vl %d vn %d\n", subject, value_prev, value, value_new); 
					}
					else {
						if (entschuldigt) {
							map2.get(name).put(subject, new Abwesenheit(value, 0));
							//System.out.printf("Ent Name %s NEUES Fach %s v %d\n", name, subject, value);
						}
						else {
							map2.get(name).put(subject, new Abwesenheit(0, value));
							//System.out.printf("UntEnt Name %s NEUES Fach %s v %d\n", name, subject, value);
						}
					}		
				}
				else {
					if ( value > 0 ) {
						if (entschuldigt) {
							map.put(name, new Abwesenheit(value, 0));
							map2.put(name, new TreeMap<String, Abwesenheit>());
							map2.get(name).put(subject, new Abwesenheit(value, 0));
							//System.out.printf("Ent NEUER Name %s NEUES Fach %s v %d\n", name, subject, value);
						}
						else {
							map.put(name, new Abwesenheit(0, value));
							map2.put(name, new TreeMap<String, Abwesenheit>());
							map2.get(name).put(subject, new Abwesenheit(0, value));
							//System.out.printf("UnEnt NEUER Name %s NEUES Fach %s v %d\n", name, subject, value);
						}	
					}
				}
				if (map3.containsKey(subject)) { 
					if (map3.get(subject).containsKey(name)) {
						if (entschuldigt) {
							int value_prev = map3.get(subject).get(name).getEntAbw();	
							int value_new =  value_prev + value;
							map3.get(subject).get(name).setEntAbw(value_new);
							//System.out.printf("Ent Fach %s Name %s vp %d  vl %d vn %d\n", subject, name, value_prev, value, value_new); 
						}
						else {
							int value_prev = map3.get(subject).get(name).getUnEntAbw();	
							int value_new =  value_prev + value;
							map3.get(subject).get(name).setUnEntAbw(value_new);
							//System.out.printf("UnEnt Fach %s Name %s vp %d  vl %d vn %d\n", subject, name, value_prev, value, value_new); 
						}			
					}
					else {
						if (entschuldigt) {
							map3.get(subject).put(name, new Abwesenheit(value, 0));
							//System.out.printf("Ent Fach %s NEUER Name %s v %d\n", subject, name, value);
						}
						else {
							map3.get(subject).put(name, new Abwesenheit(0, value));
							//System.out.printf("UntEnt Fach %s NEUER Name %s v %d\n", subject, name, value);
						}
					}		
				}
				else {
					if ( value > 0 ) {
						if (entschuldigt) {
							map3.put(subject, new TreeMap<String, Abwesenheit>());
							map3.get(subject).put(name, new Abwesenheit(value, 0));
							//System.out.printf("Ent NEUES Fach %s NEUER Name %s v %d\n", subject, name, value);
						}
						else {
							map3.put(subject, new TreeMap<String, Abwesenheit>());
							map3.get(subject).put(name, new Abwesenheit(0, value));
							//System.out.printf("UnEnt NEUES Fach %s NEUER Name %s v %d\n", subject, name, value);
						}	
					}
				}
				i++;
			}
			inp_aps.close();
			showInformationAlert(aps_str + "\nwurde eingelesen.");
			//System.out.println("\nAnzahl eingelesener Zeilen der AbsencePerStudent-Datei: " + i);
			aps_File_imported = true;
			aps_File_mustBeRead = false;
			
		} catch (FileNotFoundException ex) {
			System.err.println(ex);
			showErrorAlert("Die AbsencePerStudent-Datei " + aps_str + " konnte nicht eingelesen werden\n" + ex);
			map.clear();
			map2.clear();
			map3.clear();
		} catch (NumberFormatException ex) {
			System.err.println(ex);
			showErrorAlert("Die AbsencePerStudent-Datei " + aps_str + " konnte nicht eingelesen werden\n" + ex);
			map.clear();
			map2.clear();
			map3.clear();
		}	
	}
	
	private void erstelleFehlzeitenTabelle(Stage stage) {
		
		if(tM_File_imported == false) {
			showInformationAlert("Es wurde noch keine TeachingMethods-Datei eingelesen.");
			return;
		}
		if(mapNrOfLessonsPerSubject.isEmpty()) {
			showInformationAlert("Die TeachingMethods-Datei muss eingelesen werden.");
			return;
		}
		if(aps_File_imported == false) {
			showInformationAlert("Es wurde noch keine AbsencePerStudent-Datei eingelesen.");
			return;
		}
		if(aps_File_mustBeRead == true) {
			showInformationAlert("Die AbsencePerStudent-Datei muss neu eingelesen werden.");
			return;
		}
		if(map.isEmpty() || map2.isEmpty() || map3.isEmpty()) {
			showInformationAlert("Die AbsencePerStudent-Datei muss eingelesen werden.");
			return;
		}
		
		String fehldatei = "Fehlzeiten_" + kurs + "_" + startDate + "-" + endDate + ".csv";
		FileChooser fileChooser = new FileChooser();
		fileChooser.getExtensionFilters().add(new ExtensionFilter("CSV", "*.csv"));
		fileChooser.setInitialDirectory(new File(initialDir));
		fileChooser.setTitle("Speichere Fehlzeiten-Datei");
		fileChooser.setInitialFileName(fehldatei);
		File file = fileChooser.showSaveDialog(stage);
		String outStr = "";
		try {
			FileWriter aus_csv = new FileWriter(file, false);
			
			outStr = String.format("Fehlzeiten in der Klasse %s von\ndem %s bis zum %s\n\n", kurs, startDate, endDate);		
			aus_csv.write(outStr);		

			outStr = String.format("Die gesamte Anzahl der\nstattgefundenen Unterrichtsstunden (U-Std)\nin der Klasse %s von\ndem %s bis zum %s\nbetr\u00E4gt: %d U-Std\n\n",
					kurs, startDate, endDate, nrOfLessons);		
			aus_csv.write(outStr);
			
			outStr = String.format("\nAnzahl der stattgefundenen U-Std pro Fach:\n");
			aus_csv.write(outStr);			
			
			outStr = String.format("Fach\u003BU-Std\n");
			aus_csv.write(outStr);
			
			for (Map.Entry<String, Integer> mapIt : mapNrOfLessonsPerSubject.entrySet()) {
				outStr = String.format("%s\u003B%d\n", mapIt.getKey(), mapIt.getValue());
				aus_csv.write(outStr);
			}  
			outStr = String.format("\n");
			aus_csv.write(outStr);
			
			outStr = String.format("Sch\u00FCler\u002Ainnen mit Fehlzeiten\n");
			aus_csv.write(outStr);
			
			outStr = String.format("Name\u003BFehlzeit in U-Std\u003BFehlzeit in Prozent" +
						   "\u003Bunentsch. Fz. in U-Std\u003Bunentsch. Fz. in Prozent\n");
			aus_csv.write(outStr);
			
			for (String key : map.keySet()) {
				int efz = map.get(key).getEntAbw();
				int ufz = map.get(key).getUnEntAbw();
				int fz_ges = efz + ufz;
				double efzh = (double)efz / 45.0;
				double ufzh = (double)ufz / 45.0;
				double fzh_ges = (double)fz_ges / 45.0;
				try {
					double efzp = efzh * 100.0 / (double)nrOfLessons; 
					double ufzp = ufzh * 100.0 / (double)nrOfLessons;
					double fzp_ges = fzh_ges * 100.0 / (double)nrOfLessons;
					//outStr = String.format("%s\u003B%.2f\u003B%.2f\u003B%.2f\u003B%.2f\n", key, fzh_ges, fzp_ges, ufzh, ufzp);
					String values = String.format("%.2f\u003B%.2f\u003B%.2f\u003B%.2f", fzh_ges, fzp_ges, ufzh, ufzp);
					values = values.replace('.', ',');
					outStr = String.format("%s\u003B%s\n", key, values);
					aus_csv.write(outStr);
				} 
				catch (ArithmeticException e) {
					System.out.println("nrOfLessons " + nrOfLessons + " " + e.getMessage());
				}
			}
			int nrLPSub = 0;
			outStr = String.format("\n");
			aus_csv.write(outStr);
			
			outStr = String.format("Sch\u00FCler\u002Ainnen mit Fehlzeiten pro Fach\n");
			aus_csv.write(outStr);
			
			outStr = String.format("Name\u003BFach\u003BFehlzeit in U-Std\u003BFehlzeit in Prozent" +
						   "\u003Bunentsch. Fz. in U-Std\u003Bunentsch. Fz. in Prozent\n");
			aus_csv.write(outStr);
			
			for (Map.Entry<String, TreeMap<String, Abwesenheit>> map2It : map2.entrySet()) {
				String name = map2It.getKey();
				for (Map.Entry<String, Abwesenheit> mapIt : map2It.getValue().entrySet()) {
					String subject = mapIt.getKey();
					int efz = mapIt.getValue().getEntAbw();
					int ufz = mapIt.getValue().getUnEntAbw();
					int fz_ges = efz + ufz;
					if (mapNrOfLessonsPerSubject.containsKey(subject)) {
						nrLPSub = mapNrOfLessonsPerSubject.get(subject);
						double efzh = (double)efz / 45.0;
						double ufzh = (double)ufz / 45.0;
						double fzh_ges = (double)fz_ges / 45.0;
						try {
							double efzp = efzh * 100.0 / (double)nrLPSub; 
							double ufzp = ufzh * 100.0 / (double)nrLPSub;
							double fzp_ges = fzh_ges * 100.0 / (double)nrLPSub;
							//outStr = String.format("%s\u003B%s\u003B%.2f\u003B%.2f\u003B%.2f\u003B%.2f\n", name, subject, fzh_ges, fzp_ges, ufzh, ufzp);
							String values = String.format("%.2f\u003B%.2f\u003B%.2f\u003B%.2f", fzh_ges, fzp_ges, ufzh, ufzp);
							values = values.replace('.', ',');
							outStr = String.format("%s\u003B%s\u003B%s\n", name, subject, values);
							aus_csv.write(outStr);
						} 
						catch (ArithmeticException e) {
							System.out.println("nrLPub " + nrLPSub + " " + e.getMessage());
						}
					}
				}
			}  
			outStr = String.format("\n");
			aus_csv.write(outStr);
			
			outStr = String.format("Sch\u00FCler\u002Ainnen mit Fehlzeiten pro Fach\n");
			aus_csv.write(outStr);
			
			outStr = String.format("Fach\u003BName\u003BFehlzeit in U-Std\u003BFehlzeit in Prozent" +
						   "\u003Bunentsch. Fz. in U-Std\u003Bunentsch. Fz. in Prozent\n");
			aus_csv.write(outStr);
			
			for (Map.Entry<String, TreeMap<String, Abwesenheit>> map3It : map3.entrySet()) {
				String subject = map3It.getKey();
				for (Map.Entry<String, Abwesenheit> mapIt : map3It.getValue().entrySet()) {
					String name = mapIt.getKey();
					int efz = mapIt.getValue().getEntAbw();
					int ufz = mapIt.getValue().getUnEntAbw();
					int fz_ges = efz + ufz;
					if (mapNrOfLessonsPerSubject.containsKey(subject)) {
						nrLPSub = mapNrOfLessonsPerSubject.get(subject);
						double efzh = (double)efz / 45.0;
						double ufzh = (double)ufz / 45.0;
						double fzh_ges = (double)fz_ges / 45.0;
						try {
							double efzp = efzh * 100.0 / (double)nrLPSub; 
							double ufzp = ufzh * 100.0 / (double)nrLPSub;
							double fzp_ges = fzh_ges * 100.0 / (double)nrLPSub;
							//outStr = String.format("%s\u003B%s\u003B%.2f\u003B%.2f\u003B%.2f\u003B%.2f\n", subject, name, fzh_ges, fzp_ges, ufzh, ufzp);
							String values = String.format("%.2f\u003B%.2f\u003B%.2f\u003B%.2f", fzh_ges, fzp_ges, ufzh, ufzp);
							values = values.replace('.', ',');
							outStr = String.format("%s\u003B%s\u003B%s\n", subject, name, values);
							aus_csv.write(outStr);
						} 
						catch (ArithmeticException e) {
							System.out.println("nrLPub " + nrLPSub + " " + e.getMessage());
						}
					}
				}
			} 
			aus_csv.close();
			showInformationAlert("Die Fehlzeiten-Tabellen wurden als CSV-Datei gespeichert.");
		} catch (FileNotFoundException ex) {
			System.err.println(ex);
			showErrorAlert("Der Prozess kann nicht auf die Datei\n" + fehldatei + " zugreifen,\nda sie von einem anderen Prozess verwendet wird");
		} catch (IOException ex) {
			System.err.println(ex);
			showErrorAlert("Der Prozess kann nicht auf die Datei\n" + fehldatei + " zugreifen,\nda sie von einem anderen Prozess verwendet wird");
		}
	}
	
	private void createPDF(Stage stage) {
		
		if(tM_File_imported == false) {
			showInformationAlert("Es wurde noch keine TeachingMethods-Datei eingelesen.");
			return;
		}
		if(mapNrOfLessonsPerSubject.isEmpty()) {
			showInformationAlert("Die TeachingMethods-Datei muss eingelesen werden.");
			return;
		}
		if(aps_File_imported == false) {
			showInformationAlert("Es wurde noch keine AbsencePerStudent-Datei eingelesen.");
			return;
		}
		if(aps_File_mustBeRead == true) {
			showInformationAlert("Die AbsencePerStudent-Datei muss neu eingelesen werden.");
			return;
		}
		if(map.isEmpty() || map2.isEmpty() || map3.isEmpty()) {
			showInformationAlert("Die AbsencePerStudent-Datei muss eingelesen werden.");
			return;
		}
		
		String fehldatei = "Fehlzeiten_" + kurs + "_" + startDate + "-" + endDate + ".pdf";
		
		Document document = new Document(PageSize.A4, 50, 50, 50, 50);
		FileOutputStream fileOS_doc = null;
		
		FileChooser fileChooser = new FileChooser();
		fileChooser.getExtensionFilters().add(new ExtensionFilter("PDF", "*.pdf"));
		fileChooser.setInitialDirectory(new File(initialDir));
		fileChooser.setTitle("Erzeuge Fehlzeiten-PDF-Datei");
		fileChooser.setInitialFileName(fehldatei);
		File file = fileChooser.showSaveDialog(stage);
		String file_str = file.getName();
		String outStr = "";
		try {
			
			fileOS_doc = new FileOutputStream(file);
			outStr = String.format("Fehlzeiten in der Klasse %s von\ndem %s bis zum %s\n\n", kurs, startDate, endDate);		
			PdfWriter.getInstance(document, fileOS_doc).setInitialLeading(10);
            
			//LineSeparator line = new LineSeparator(1, 100, null, Element.ALIGN_CENTER, -2);
			Font helv_12 = new Font(Font.HELVETICA, 12, Font.NORMAL, new Color(0, 0, 0));
			Font helv_13 = new Font(Font.HELVETICA, 13, Font.NORMAL, new Color(0, 0, 0));
			Font helv_13_b = new Font(Font.HELVETICA, 13, Font.BOLD, new Color(0, 0, 0));
			Font helv_14_b = new Font(Font.HELVETICA, 14, Font.BOLD, new Color(0, 0, 0));
			Font helv_15_b = new Font(Font.HELVETICA, 15, Font.BOLD, new Color(0, 0, 0));
			Font helv_16_b = new Font(Font.HELVETICA, 16, Font.BOLD, new Color(0, 0, 0));
			
			Font helv_12_blue = new Font(Font.HELVETICA, 12, Font.BOLD, new Color(0, 0, 255));
			Font helv_12_red = new Font(Font.HELVETICA, 12, Font.NORMAL, new Color(255, 0, 0));
			Font helv_12_red_b = new Font(Font.HELVETICA, 12, Font.BOLD, new Color(255, 0, 0));
			Font helv_14_blue = new Font(Font.HELVETICA, 14, Font.BOLD, new Color(0, 0, 255));
			Font helv_12_green = new Font(Font.HELVETICA, 12, Font.BOLD, new Color(60, 195, 0));
			Font helv_11_black = new Font(Font.HELVETICA, 11);
			Font helv_12_black = new Font(Font.HELVETICA, 12);
			Font helv_12_black_b = new Font(Font.HELVETICA, 12, Font.BOLD);
			
			Font timesRoman_11 = new Font(Font.TIMES_ROMAN, 11);
			Font timesRoman_12 = new Font(Font.TIMES_ROMAN, 12);
			
			HeaderFooter footer = new HeaderFooter(new Phrase("", timesRoman_11), true);
            footer.setBorder(Rectangle.NO_BORDER);
            footer.setAlignment(Element.ALIGN_CENTER);
            document.setFooter(footer);
			
			outStr = String.format("Fehlzeiten in der Klasse %s vom %s bis zum %s", kurs, startDate, endDate);
			HeaderFooter head = new HeaderFooter(new Phrase(outStr, timesRoman_11), false);
            head.setAlignment(Element.ALIGN_LEFT);
            document.setHeader(head);
			
            document.open();
			
			outStr = String.format("Fehlzeiten in der Klasse %s vom %s bis zum %s\n", kurs, startDate, endDate);	
			Paragraph par = new Paragraph(outStr, helv_15_b);
            document.add(par);	
			
			Anchor webUntis = new Anchor("WebUntis", helv_12_blue);
			webUntis.setReference("https://webuntis.com/");
			
			outStr = String.format("Auswertung der von ");	
			par = new Paragraph(outStr, helv_12);
			par.add(webUntis);
			par.add(new Chunk(" unter Klassenbuch -> Berichte heruntergeladenen Dateien:\n", helv_12));
            document.add(par);
			
			List list_webUntisReports = new List(List.ORDERED, List.ALPHABETICAL);
			list_webUntisReports.setIndentationLeft(100);
			ListItem item = new ListItem("  " + aps_str, helv_12);
			item.setListSymbol(new Chunk("\u25CB "));
			list_webUntisReports.add(item);
			item = new ListItem("  " + tM_str, helv_12);
			item.setListSymbol(new Chunk("\u25CB "));
			list_webUntisReports.add(item);
			document.add(list_webUntisReports);
			
			document.add(Chunk.NEWLINE);
			
			String header = "Anzahl der stattgefundenen Unterrichtsstunden";
			par = new Paragraph(header, helv_13_b);
			par.setAlignment(Element.ALIGN_CENTER);
            document.add(par);
			
			outStr = String.format("Die gesamte Anzahl der stattgefundenen Unterrichtsstunden (U-Std) in der Klasse %s vom %s bis zum %s betr\u00E4gt: %d U-Std\n\n",
					kurs, startDate, endDate, nrOfLessons);		
			par = new Paragraph(outStr, helv_12);
            document.add(par);
			
			header = "Anzahl der stattgefundenen U-Std pro Fach";
			par = new Paragraph(header, helv_13_b);
			par.setAlignment(Element.ALIGN_CENTER);
            document.add(par);
			
			outStr = String.format("In folgender Tabelle sind die stattgefundenen Unterrichtsstunden pro Fach dargestellt:\n");		
			par = new Paragraph(outStr, helv_12);
            document.add(par);
			PdfPTable table1 = new PdfPTable(2);
			table1.setSpacingBefore(10);
			table1.setSpacingAfter(10);
			
			table1.setTotalWidth(250);
			table1.setLockedWidth(true);
			table1.setWidths(new float[]{1,1});
			table1.getDefaultCell().setBorderColor(new Color(0, 0, 255));
            table1.getDefaultCell().setPadding(5);
			//table1.getDefaultCell().setFixedHeight(20);
            table1.getDefaultCell().setHorizontalAlignment(Element.ALIGN_CENTER);
			table1.getDefaultCell().setVerticalAlignment(Element.ALIGN_CENTER);
			
			PdfPCell cellTab1 = new PdfPCell(new Paragraph("Fach", helv_12_black_b));
			//cellTab1.setFixedHeight(20);
			cellTab1.setMinimumHeight(20);
			cellTab1.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab1.setVerticalAlignment(Element.ALIGN_CENTER);
			table1.addCell(cellTab1);
			
			cellTab1 = new PdfPCell(new Paragraph("U-Std", helv_12_black_b));
			//cellTab1.setFixedHeight(20);
			cellTab1.setMinimumHeight(20);
			cellTab1.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab1.setVerticalAlignment(Element.ALIGN_CENTER);
			table1.addCell(cellTab1);
			
			PdfPCell cellFU;
			for (Map.Entry<String, Integer> mapIt : mapNrOfLessonsPerSubject.entrySet()) {
				cellFU = new PdfPCell(new Paragraph(mapIt.getKey(), helv_12_black));
				cellFU.setPaddingTop(5);
				cellFU.setPaddingBottom(5);
				cellFU.setHorizontalAlignment(Element.ALIGN_CENTER);
				cellFU.setVerticalAlignment(Element.ALIGN_CENTER);
				table1.addCell(cellFU);
				
				cellFU = new PdfPCell(new Paragraph(String.valueOf(mapIt.getValue()).replace(".",","), helv_12_black));
				cellFU.setPaddingTop(5);
				cellFU.setPaddingBottom(5);
				cellFU.setHorizontalAlignment(Element.ALIGN_CENTER);
				cellFU.setVerticalAlignment(Element.ALIGN_CENTER);
				table1.addCell(cellFU);
			} 
			
			document.add(table1);

			outStr = String.format("\n");
			par = new Paragraph(outStr, helv_12_black);
            document.add(par);
			
			header = "Sch\u00FCler\u002Ainnen mit Fehlzeiten \u00FCber alle F\u00E4cher";
			par = new Paragraph(header, helv_13_b);
			par.setAlignment(Element.ALIGN_CENTER);
            document.add(par);
			
			outStr = String.format("In folgender Tabelle stehen die alphabetisch nach Nachnamen sortierten Fehlzeiten aller Sch\u00FCler\u002Ainnen in Bezug auf alle stattgefundenen Unterrichtsstunden:\n");		
			par = new Paragraph(outStr, helv_12);
            document.add(par);
			
			PdfPTable table2 = new PdfPTable(5);
			table2.setSpacingBefore(10);
			table2.setSpacingAfter(10);
			
			table2.setTotalWidth(490);
			table2.setLockedWidth(true);
			table2.setWidths(new float[]{2,1,1,1,1});
			table2.getDefaultCell().setBorderColor(new Color(0, 0, 255));
            table2.getDefaultCell().setPadding(7);
            table2.getDefaultCell().setHorizontalAlignment(Element.ALIGN_CENTER);
			table2.getDefaultCell().setVerticalAlignment(Element.ALIGN_CENTER);
			
			PdfPCell cellTab2 = new PdfPCell(new Paragraph("Name", helv_12_black_b));
			cellTab2.setPaddingTop(10);
			cellTab2.setPaddingBottom(5);
			cellTab2.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab2.setVerticalAlignment(Element.ALIGN_CENTER);
			table2.addCell(cellTab2);
			
			cellTab2 = new PdfPCell(new Paragraph("Fehlzeit\nin U-Std", helv_12_black_b));
			cellTab2.setPaddingTop(5);
			cellTab2.setPaddingBottom(5);
			cellTab2.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab2.setVerticalAlignment(Element.ALIGN_CENTER);
			table2.addCell(cellTab2);
			
			cellTab2 = new PdfPCell(new Paragraph("Fehlzeit\nin Prozent", helv_12_black_b));
			cellTab2.setPaddingTop(5);
			cellTab2.setPaddingBottom(5);
			cellTab2.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab2.setVerticalAlignment(Element.ALIGN_CENTER);
			table2.addCell(cellTab2);
			
			cellTab2 = new PdfPCell(new Paragraph("unentsch. Fehlzeit\nin U-Std", helv_12_black_b));
			cellTab2.setPaddingTop(5);
			cellTab2.setPaddingBottom(5);
			cellTab2.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab2.setVerticalAlignment(Element.ALIGN_CENTER);
			table2.addCell(cellTab2);
			
			cellTab2 = new PdfPCell(new Paragraph("unentsch. Fehlzeit\nin Prozent", helv_12_black_b));
			cellTab2.setPaddingTop(5);
			cellTab2.setPaddingBottom(5);
			cellTab2.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab2.setVerticalAlignment(Element.ALIGN_CENTER);
			table2.addCell(cellTab2);
			
			PdfPCell cellFZ;
			int i = 1;
			for (String key : map.keySet()) {
				int efz = map.get(key).getEntAbw();
				int ufz = map.get(key).getUnEntAbw();
				int fz_ges = efz + ufz;
				double efzh = (double)efz / 45.0;
				double ufzh = (double)ufz / 45.0;
				double fzh_ges = (double)fz_ges / 45.0;
				try {
					double efzp = efzh * 100.0 / (double)nrOfLessons; 
					double ufzp = ufzh * 100.0 / (double)nrOfLessons;
					double fzp_ges = fzh_ges * 100.0 / (double)nrOfLessons;
					
					cellFZ = new PdfPCell(new Paragraph(key, helv_12_black));
					cellFZ.setPaddingTop(5);
					cellFZ.setPaddingBottom(5);
					cellFZ.setHorizontalAlignment(Element.ALIGN_LEFT);
					if (i % 2 == 1) {
						cellFZ.setGrayFill(0.9f);
					}
					table2.addCell(cellFZ);
					
					String fzh_ges_str = String.valueOf(round(fzh_ges, 2)).replace(".",",");
					cellFZ = new PdfPCell(new Paragraph(fzh_ges_str, helv_12_black));
					cellFZ.setPaddingTop(5);
					cellFZ.setPaddingBottom(5);
					cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
					if (i % 2 == 1) {
						cellFZ.setGrayFill(0.9f);
					}
					table2.addCell(cellFZ);
					
					String fzp_ges_str = String.valueOf(round(fzp_ges, 2)).replace(".",",") + " \u0025";
					cellFZ = new PdfPCell(new Paragraph(fzp_ges_str, helv_12_black));
					cellFZ.setCellEvent(new PercentCell(fzp_ges));
					cellFZ.setPaddingTop(5);
					cellFZ.setPaddingBottom(5);
					cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
					if (i % 2 == 1) {
						cellFZ.setGrayFill(0.9f);
					}
					table2.addCell(cellFZ);
					
					String ufzh_str = String.valueOf(round(ufzh, 2)).replace(".",",");
					cellFZ = new PdfPCell(new Paragraph(ufzh_str, helv_12_black));
					cellFZ.setPaddingTop(5);
					cellFZ.setPaddingBottom(5);
					cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
					if (i % 2 == 1) {
						cellFZ.setGrayFill(0.9f);
					}
					table2.addCell(cellFZ);
					
					String ufzp_str = String.valueOf(round(ufzp, 2)).replace(".",",") + " \u0025";
					cellFZ = new PdfPCell(new Paragraph(ufzp_str, helv_12_black));
					cellFZ.setCellEvent(new PercentCell(ufzp));
					cellFZ.setPaddingTop(5);
					cellFZ.setPaddingBottom(5);
					cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
					if (i % 2 == 1) {
						cellFZ.setGrayFill(0.9f);
					}
					table2.addCell(cellFZ);
				} 
				catch (ArithmeticException e) {
					System.out.println("nrOfLessons " + nrOfLessons + " " + e.getMessage());
				}
				i++;
			}
			
			document.add(table2);
			
			document.add(Chunk.NEWLINE);
			
			header = "Sch\u00FCler\u002Ainnen mit Fehlzeiten pro Fach";
			par = new Paragraph(header, helv_13_b);
			par.setAlignment(Element.ALIGN_CENTER);
            document.add(par);
			
			outStr = String.format("In folgender Tabelle stehen die alphabetisch nach Nachnamen und F\u00E4chern sortierten Fehlzeiten aller Sch\u00FCler\u002Ainnen pro Fach:\n");		
			par = new Paragraph(outStr, helv_12);
            document.add(par);
			
			PdfPTable table3 = new PdfPTable(6);
			table3.setSpacingBefore(10);
			table3.setSpacingAfter(10);
			
			table3.setTotalWidth(490);
			table3.setLockedWidth(true);
			table3.setWidths(new float[]{2,1,1,1,1,1});
			table3.getDefaultCell().setBorderColor(new Color(0, 0, 255));
            table3.getDefaultCell().setPadding(7);
            table3.getDefaultCell().setHorizontalAlignment(Element.ALIGN_CENTER);
			table3.getDefaultCell().setVerticalAlignment(Element.ALIGN_CENTER);
			
			PdfPCell cellTab3 = new PdfPCell(new Paragraph("Name", helv_12_black_b));
			cellTab3.setPaddingTop(10);
			cellTab3.setPaddingBottom(5);
			cellTab3.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab3.setVerticalAlignment(Element.ALIGN_CENTER);
			table3.addCell(cellTab3);
			
			cellTab3 = new PdfPCell(new Paragraph("Fach", helv_12_black_b));
			cellTab3.setPaddingTop(10);
			cellTab3.setPaddingBottom(5);
			cellTab3.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab3.setVerticalAlignment(Element.ALIGN_CENTER);
			table3.addCell(cellTab3);
			
			cellTab3 = new PdfPCell(new Paragraph("Fehlzeit\nin U-Std", helv_12_black_b));
			cellTab3.setPaddingTop(5);
			cellTab3.setPaddingBottom(5);
			cellTab3.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab3.setVerticalAlignment(Element.ALIGN_CENTER);
			table3.addCell(cellTab3);
			
			cellTab3 = new PdfPCell(new Paragraph("Fehlzeit\nin Prozent", helv_12_black_b));
			cellTab3.setPaddingTop(5);
			cellTab3.setPaddingBottom(5);
			cellTab3.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab3.setVerticalAlignment(Element.ALIGN_CENTER);
			table3.addCell(cellTab3);
			
			cellTab3 = new PdfPCell(new Paragraph("unentsch. Fehlzeit\nin U-Std", helv_12_black_b));
			cellTab3.setPaddingTop(5);
			cellTab3.setPaddingBottom(5);
			cellTab3.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab3.setVerticalAlignment(Element.ALIGN_CENTER);
			table3.addCell(cellTab3);
			
			cellTab3 = new PdfPCell(new Paragraph("unentsch. Fehlzeit\nin Prozent", helv_12_black_b));
			cellTab3.setPaddingTop(5);
			cellTab3.setPaddingBottom(5);
			cellTab3.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab3.setVerticalAlignment(Element.ALIGN_CENTER);
			table3.addCell(cellTab3);
			
			int nrLPSub = 0;
			i = 1;
			for (Map.Entry<String, TreeMap<String, Abwesenheit>> map2It : map2.entrySet()) {
				String name = map2It.getKey();
				for (Map.Entry<String, Abwesenheit> mapIt : map2It.getValue().entrySet()) {
					String subject = mapIt.getKey();
					int efz = mapIt.getValue().getEntAbw();
					int ufz = mapIt.getValue().getUnEntAbw();
					int fz_ges = efz + ufz;
					if (mapNrOfLessonsPerSubject.containsKey(subject)) {
						nrLPSub = mapNrOfLessonsPerSubject.get(subject);
						double efzh = (double)efz / 45.0;
						double ufzh = (double)ufz / 45.0;
						double fzh_ges = (double)fz_ges / 45.0;
						try {
							double efzp = efzh * 100.0 / (double)nrLPSub; 
							double ufzp = ufzh * 100.0 / (double)nrLPSub;
							double fzp_ges = fzh_ges * 100.0 / (double)nrLPSub;
							
							cellFZ = new PdfPCell(new Paragraph(name, helv_12_black));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_LEFT);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table3.addCell(cellFZ);
							
							cellFZ = new PdfPCell(new Paragraph(subject, helv_12_black));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table3.addCell(cellFZ);
							
							String fzh_ges_str = String.valueOf(round(fzh_ges, 2)).replace(".",",");
							cellFZ = new PdfPCell(new Paragraph(fzh_ges_str, helv_12_black));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table3.addCell(cellFZ);
							
							String fzp_ges_str = String.valueOf(round(fzp_ges, 2)).replace(".",",") + " \u0025";
							cellFZ = new PdfPCell(new Paragraph(fzp_ges_str, helv_12_black));
							cellFZ.setCellEvent(new PercentCell(fzp_ges));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table3.addCell(cellFZ);
							
							String ufzh_str = String.valueOf(round(ufzh, 2)).replace(".",",");
							cellFZ = new PdfPCell(new Paragraph(ufzh_str, helv_12_black));
							
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table3.addCell(cellFZ);
							
							String ufzp_str = String.valueOf(round(ufzp, 2)).replace(".",",") + " \u0025";
							cellFZ = new PdfPCell(new Paragraph(ufzp_str, helv_12_black));
							cellFZ.setCellEvent(new PercentCell(ufzp));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table3.addCell(cellFZ);
						} 
						catch (ArithmeticException e) {
							System.out.println("nrLPub " + nrLPSub + " " + e.getMessage());
						}
						i++;
					}
				}
			}  
			
			document.add(table3);
			
			document.add(Chunk.NEWLINE);
			
			header = "Sch\u00FCler\u002Ainnen mit Fehlzeiten pro Fach";
			par = new Paragraph(header, helv_13_b);
			par.setAlignment(Element.ALIGN_CENTER);
            document.add(par);
			
			outStr = String.format("In folgender Tabelle stehen die alphabetisch nach F\u00E4chern und Nachnamen sortierten Fehlzeiten aller Sch\u00FCler\u002Ainnen pro Fach:\n");		
			par = new Paragraph(outStr, helv_12);
            document.add(par);
			
			PdfPTable table4 = new PdfPTable(6);
			table4.setSpacingBefore(10);
			table4.setSpacingAfter(10);
			
			table4.setTotalWidth(490);
			table4.setLockedWidth(true);
			table4.setWidths(new float[]{1,2,1,1,1,1});
			table4.getDefaultCell().setBorderColor(new Color(0, 0, 255));
            table4.getDefaultCell().setPadding(7);
            table4.getDefaultCell().setHorizontalAlignment(Element.ALIGN_CENTER);
			table4.getDefaultCell().setVerticalAlignment(Element.ALIGN_CENTER);
			
			PdfPCell cellTab4 = new PdfPCell(new Paragraph("Fach", helv_12_black_b));
			cellTab4.setPaddingTop(10);
			cellTab4.setPaddingBottom(5);
			cellTab4.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab4.setVerticalAlignment(Element.ALIGN_CENTER);
			table4.addCell(cellTab4);
			
			cellTab4 = new PdfPCell(new Paragraph("Name", helv_12_black_b));
			cellTab4.setPaddingTop(10);
			cellTab4.setPaddingBottom(5);
			cellTab4.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab4.setVerticalAlignment(Element.ALIGN_CENTER);
			table4.addCell(cellTab4);
			
			cellTab4 = new PdfPCell(new Paragraph("Fehlzeit\nin U-Std", helv_12_black_b));
			cellTab4.setPaddingTop(5);
			cellTab4.setPaddingBottom(5);
			cellTab4.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab4.setVerticalAlignment(Element.ALIGN_CENTER);
			table4.addCell(cellTab4);
			
			cellTab4 = new PdfPCell(new Paragraph("Fehlzeit\nin Prozent", helv_12_black_b));
			cellTab4.setPaddingTop(5);
			cellTab4.setPaddingBottom(5);
			cellTab4.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab4.setVerticalAlignment(Element.ALIGN_CENTER);
			table4.addCell(cellTab4);
			
			cellTab4 = new PdfPCell(new Paragraph("unentsch. Fehlzeit\nin U-Std", helv_12_black_b));
			cellTab4.setPaddingTop(5);
			cellTab4.setPaddingBottom(5);
			cellTab4.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab4.setVerticalAlignment(Element.ALIGN_CENTER);
			table4.addCell(cellTab4);
			
			cellTab4 = new PdfPCell(new Paragraph("unentsch. Fehlzeit\nin Prozent", helv_12_black_b));
			cellTab4.setPaddingTop(5);
			cellTab4.setPaddingBottom(5);
			cellTab4.setHorizontalAlignment(Element.ALIGN_CENTER);
			cellTab4.setVerticalAlignment(Element.ALIGN_CENTER);
			table4.addCell(cellTab4);
			
			nrLPSub = 0;
			i = 1;
			for (Map.Entry<String, TreeMap<String, Abwesenheit>> map3It : map3.entrySet()) {
				String subject = map3It.getKey();
				for (Map.Entry<String, Abwesenheit> mapIt : map3It.getValue().entrySet()) {
					String name = mapIt.getKey();
					int efz = mapIt.getValue().getEntAbw();
					int ufz = mapIt.getValue().getUnEntAbw();
					int fz_ges = efz + ufz;
					if (mapNrOfLessonsPerSubject.containsKey(subject)) {
						nrLPSub = mapNrOfLessonsPerSubject.get(subject);
						double efzh = (double)efz / 45.0;
						double ufzh = (double)ufz / 45.0;
						double fzh_ges = (double)fz_ges / 45.0;
						try {
							double efzp = efzh * 100.0 / (double)nrLPSub; 
							double ufzp = ufzh * 100.0 / (double)nrLPSub;
							double fzp_ges = fzh_ges * 100.0 / (double)nrLPSub;
							
							cellFZ = new PdfPCell(new Paragraph(subject, helv_12_black));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table4.addCell(cellFZ);
							
							cellFZ = new PdfPCell(new Paragraph(name, helv_12_black));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_LEFT);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table4.addCell(cellFZ);
							
							String fzh_ges_str = String.valueOf(round(fzh_ges, 2)).replace(".",",");
							cellFZ = new PdfPCell(new Paragraph(fzh_ges_str, helv_12_black));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table4.addCell(cellFZ);
							
							String fzp_ges_str = String.valueOf(round(fzp_ges, 2)).replace(".",",") + " \u0025";
							cellFZ = new PdfPCell(new Paragraph(fzp_ges_str, helv_12_black));
							cellFZ.setCellEvent(new PercentCell(fzp_ges));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table4.addCell(cellFZ);
							
							String ufzh_str = String.valueOf(round(ufzh, 2)).replace(".",",");
							cellFZ = new PdfPCell(new Paragraph(ufzh_str, helv_12_black));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table4.addCell(cellFZ);
							
							String ufzp_str = String.valueOf(round(ufzp, 2)).replace(".",",") + " \u0025";
							cellFZ = new PdfPCell(new Paragraph(ufzp_str, helv_12_black));
							cellFZ.setCellEvent(new PercentCell(ufzp));
							cellFZ.setPaddingTop(5);
							cellFZ.setPaddingBottom(5);
							cellFZ.setHorizontalAlignment(Element.ALIGN_CENTER);
							if (i % 2 == 1) {
								cellFZ.setGrayFill(0.9f);
							}
							table4.addCell(cellFZ);
						} 
						catch (ArithmeticException e) {
							System.out.println("nrLPub " + nrLPSub + " " + e.getMessage());
						}
						i++;
					}
				}
			}  
			
			document.add(table4);
			
			showInformationAlert("Die Fehlzeiten-Tabellen wurden in einer PDF-Datei gespeichert.");
			
		} catch (DocumentException | IOException de) {
            System.err.println(de.getMessage());
			showErrorAlert("Der Prozess kann nicht auf die PDF-Datei zugreifen,\nda sie von einem anderen Prozess verwendet wird.");
		}
		try {
			if (document != null) {
				document.close();
			}
			if (fileOS_doc != null) {
				fileOS_doc.close();
			}
		} catch (DocumentException | IOException ex) {
			System.err.println(ex);
		}		
	}
}
