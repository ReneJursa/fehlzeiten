package lib;

import org.openpdf.text.Rectangle;
import org.openpdf.text.pdf.PdfPCell;
import org.openpdf.text.pdf.PdfPCellEvent;
import org.openpdf.text.pdf.PdfContentByte;
import org.openpdf.text.pdf.PdfPTable;

public class PercentCell implements PdfPCellEvent {
	
	private double value;
	
	public PercentCell(double value) {
		this.value = value;
	}
	
	public void cellLayout(PdfPCell cell, Rectangle position, PdfContentByte[] canvases) {
		
		PdfContentByte pdfContentByte = canvases[PdfPTable.BACKGROUNDCANVAS];
		
		pdfContentByte.saveState();
		
		if (value < 10.0) {
			pdfContentByte.setRGBColorFill(0, 238, 0);
		}
		else if (value < 15.0) {
			pdfContentByte.setRGBColorFill(238, 238, 0);
		} 
		else if (value < 20.0) {
			pdfContentByte.setRGBColorFill(255, 165, 0);
		} 
		else {
			pdfContentByte.setRGBColorFill(255, 0, 0);
		}
		
		pdfContentByte.rectangle(position.getLeft(), position.getBottom(), position.getWidth() * (float) (value / 100.0), position.getHeight());
		
		pdfContentByte.fill();
		
		pdfContentByte.restoreState();
	}
}