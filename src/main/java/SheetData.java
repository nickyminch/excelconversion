import org.apache.maven.plugin.logging.Log;

public class SheetData {
	private String name;
	private int orientation;
	private final Log log;
	public SheetData(String name, int orientation, Log log) {
		super();
		this.name = name;
		this.log = log;
		getLog().info("orientation="+orientation);
		this.orientation = orientation;
		if(orientation!=ExcelToTextFile.HORISONTAL && orientation!=ExcelToTextFile.VERTICAL) {
			throw new RuntimeException("Unknown sheet orientation");
		}
	}
	public String getName() {
		return name;
	}
	public int getOrientation() {
//		return orientation;
		return 1;
	}
	public Log getLog() {
		return log;
	}
}
