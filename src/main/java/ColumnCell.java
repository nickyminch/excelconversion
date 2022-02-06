
public class ColumnCell {
	private String cellContents;
	
	public ColumnCell(String contents) {
		
	}
	
	public String getCellContents() {
		return cellContents;
	}

	public void setCellContents(String cellContent) {
		this.cellContents = cellContent;
	}
	
	public int getCellWidth() {
		return cellContents.length();
	}
}
