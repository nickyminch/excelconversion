import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.maven.plugin.logging.Log;
import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.formula.DataValidationEvaluator;
import org.apache.poi.ss.formula.EvaluationConditionalFormatRule;
import org.apache.poi.ss.formula.WorkbookEvaluatorProvider;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToTextFile {
    private static final String SOURCE_EXTENSION = "xlsx";
    private static final String TARGET_EXTENSION = "txt";
    private static final Predicate<String> containsTarget = path -> path.contains("\\target");
    private static final Predicate<String> isExcelFile = path -> path.endsWith("xlsx");

    private Map<String, List<List<String>>> sheetTable = new HashMap<String, List<List<String>>>();
    
    public static final int HORISONTAL = 0;
    public static final int VERTICAL = 1;

    private final Log log;
    private final String searchDirectory;

    public ExcelToTextFile(Log log, String searchDirectory) {
        this.searchDirectory = searchDirectory;
        this.log = log;
    }

    public void generateTextFilesFromExcelFile() {
        try {
            getAllExcelFiles().forEach(this::convertExcelToTextFile);
        } catch (IOException e) {
            log.error("There is no xlsx file to be converted to text file", e);
        }
    }

    private void convertExcelToTextFile(String pathToExcel) {
    	getLog().debug(pathToExcel);
        StringBuilder fileContent = new StringBuilder();

        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(pathToExcel))) {
            File file = new File(createTextPath(pathToExcel));
            for (Sheet sheet : workbook) {
            	int orientation = getSheetOrientation(sheet);
            	processSheet(sheet, orientation, fileContent);
                fileContent.append(System.lineSeparator());
            }
            Files.write(file.toPath(), String.valueOf(fileContent).getBytes(StandardCharsets.UTF_8));
        } catch (Exception e) {
            log.error(printFileEmptyOrDamagedMessage(pathToExcel), e);
        }
    }
    
    private void processSheet(Sheet sheet, int orientation, StringBuilder fileContent) {
        getLog().debug("in 2>>>>");
    	
		appendSheetContent(sheet);
    	if(orientation==VERTICAL) {
    		swapArrays(sheet);
    	}
    	
    	removeEmptyColumnsOrRows(sheet);
        
        getLog().debug("in 3>>>>");
    	
        appendSheetName(sheet.getSheetName(), fileContent);
        Map<Integer, Integer> alignment = null;

        alignment = allign(sheet.getSheetName());

       	writeToFile(sheet.getSheetName(), fileContent, alignment);
        
    }

    private String createTextPath(String fileName) {
        StringBuilder filePath = new StringBuilder();
        int extensionIndex = fileName.lastIndexOf(SOURCE_EXTENSION);

        if (extensionIndex > -1) {
            filePath.append(fileName, 0, extensionIndex)
                    .append(TARGET_EXTENSION);
        } else {
            filePath.append(fileName);
        }
        return filePath.toString();
    }

    private void appendSheetName(String sheetName, StringBuilder fileContent) {
        fileContent.append("============").append(sheetName).append("========================\n");
    }

    private void appendSheetContent(Sheet sheet) {
        List<List<String>> sheetTableLocal = new LinkedList<>();
        sheetTable.put(sheet.getSheetName(), sheetTableLocal);
        Map<Integer, Boolean> columnLengths = new HashMap<Integer, Boolean>();
        Integer columnIndex = -1;
        Integer rowIndex = -1;
        int columnSize = -1;

        for (Row row : sheet) {
            // For some rows, getLastCellNum returns -1. These rows must be igonred
            columnIndex = -1;
            
            if (row.getLastCellNum() > 0) {
                List<String> columns = new LinkedList<>();
            	rowIndex++;

                for (Cell cell : row) {
                    String cellContent;
                	columnIndex++;
                	
                    if (!CellType._NONE.equals(cell.getCellType()) && !CellType.BLANK.equals(cell.getCellType())) {
                        cellContent = getCellValueAndCharacteristics(cell, sheet.getWorkbook().getFontAt(cell.getCellStyle().getFontIndexAsInt()));
                    } else {
                        cellContent = "";
                    }

                    cellContent = cellContent.trim();
					if(cellContent.isEmpty()) {
                    	if(columnIndex>=1) {
                    		columnLengths.put(rowIndex, Boolean.TRUE);
                    	}
                    }
                    int diff = cell.getColumnIndex()-columns.size();
                    for(int i=0;i<diff;i++) {
                    	columns.add("");
                    }

                   	columns.add(cellContent);
                }

            	if(rowIndex==0) {
            		columnSize = columns.size();
            	}
            	
                int diff = columnSize-row.getLastCellNum();
                for(int i=0;i<diff;i++) {
                	columns.add("");
                }
                if(columns.size()>0) {
                	sheetTableLocal.add(columns);
                }
            }
        }
    }
    
    private int getSheetOrientation(Sheet sheet) {
        Integer columnIndex = -1;
        Integer rowIndex = -1;
        
        int lockedCountHor = -1;
        int lockedCountVert = -1;

        getLog().debug("");
        getLog().debug("sheetName="+sheet.getSheetName());
        getLog().debug("");

        for (Row row : sheet) {
            // For some rows, getLastCellNum returns -1. These rows must be igonred
            columnIndex = -1;
            
            if (row.getLastCellNum() > 0) {
            	rowIndex++;
                for (Cell cell : row) {
                	columnIndex++;
                	boolean locked = cell.getCellStyle().getLocked();
                	
                	if(rowIndex==0) {
                		if(locked) {
                    		lockedCountHor++;
                		}
                	}
                	if(columnIndex==0) {
                		if(locked) {
                    		lockedCountVert++;
                		}
                	}
                }
            }
        }
        if(lockedCountHor==columnIndex) {
			getLog().debug("Returning VERTICAL");
			return VERTICAL;
        }else if(lockedCountVert==rowIndex) {
			getLog().debug("Returning HORISONTAL");
			return HORISONTAL;
        }
        return HORISONTAL;
    }

    private Map<Integer, Integer>  allign(String sheetName) {
    	Map<Integer, Integer> sheetColLengthLocal = new HashMap<Integer, Integer>();
    	List<List<String>> sheetTableLocal = sheetTable.get(sheetName);
        Integer columnIndex = -1;

        for (List<String> row : sheetTableLocal) {
            ListIterator<String> colIterator = row.listIterator();
            columnIndex=-1;

            while (colIterator.hasNext()) {
                columnIndex++;

                Integer cellLength = colIterator.next().length();
                
                Integer oldLength = sheetColLengthLocal.get(columnIndex);
                if(oldLength==null) {
                	oldLength = 0;
                }
                cellLength = Math.max(cellLength, oldLength);
                sheetColLengthLocal.put(columnIndex, cellLength);
            }
        }
        return sheetColLengthLocal;
    }

    private void writeToFile(String sheetName, StringBuilder fileContent, Map<Integer, Integer> sheetColLength) {
    	List<List<String>> sheetTableLocal = sheetTable.get(sheetName);
        Integer columnIndex = -1;
        
        Integer lastRowLength = 0;
        lastRowLength = 0;
        for (List<String> row : sheetTableLocal) {
            ListIterator<String> colIterator = row.listIterator();
            columnIndex=-1;
            
            while (colIterator.hasNext()) {
                columnIndex++;

                String cellContent = colIterator.next();
                
           		lastRowLength = sheetColLength.get(columnIndex);

                String formatString = "%-" + lastRowLength + "s";
                if(lastRowLength>0) {
	                String formattedCellContent = String.format(formatString, cellContent);
	
	                fileContent.append(formattedCellContent);
	                fileContent.append(" | ");
                }
            }

           	fileContent.append(System.lineSeparator());
        }
    }

    private String getCellValueAndCharacteristics(Cell cell, Font font) {
        StringBuilder cellContent = new StringBuilder();

        switch (cell.getCellType()) {
            case NUMERIC:
                cellContent.append(cell.getNumericCellValue());
                break;
            case STRING:
                cellContent.append(cell.getStringCellValue());
                break;
            case FORMULA:
                cellContent.append(cell.getCellFormula());
                break;
            case BOOLEAN:
                cellContent.append(cell.getBooleanCellValue());
                break;
            case ERROR:
                cellContent.append(cell.getErrorCellValue());
                break;
            default:
                break;
        }

        appendCellAllowedValues(cell, cellContent);
        appendCellColor(cell, cellContent);
        appendCellComment(cell, cellContent);
        appendUnderline(font, cellContent);
        appendBold(font, cellContent);

        return cellContent.toString();
    }

    private void appendCellAllowedValues(Cell cell, StringBuilder fileContent) {
        if (!cell.getSheet().getDataValidations().isEmpty()) {
            WorkbookEvaluatorProvider workbookEvaluatorProvider = (WorkbookEvaluatorProvider) cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            DataValidationEvaluator dataValidationEvaluator = new DataValidationEvaluator(cell.getSheet().getWorkbook(), workbookEvaluatorProvider);
            DataValidation cellDataValidation = dataValidationEvaluator.getValidationForCell(new CellReference(cell));
            if (cellDataValidation != null) {
                String[] listValues = cellDataValidation.getValidationConstraint().getExplicitListValues();
                fileContent.append("[");
                fileContent.append(String.join(", ", listValues));
                fileContent.append("] ");
            }
        }
    }

    private void appendUnderline(Font font, StringBuilder fileContent) {
        byte underline = font.getUnderline();
        if (underline == 1) {
            fileContent.append("Underline").append(" ");
        }
    }

    private void appendBold(Font font, StringBuilder fileContent) {
        if (font.getBold()) {
            fileContent.append("Bold").append(" ");
        }
    }

    private void appendCellColor(Cell cell, StringBuilder fileContent) {
        // (Foreground) Cell Color not set by Conditional Formatting
        XSSFColor foreColor = (XSSFColor) cell.getCellStyle().getFillForegroundColorColor();
        if (foreColor != null) {
            fileContent.append(" ").append("#").append(foreColor.getARGBHex()).append(" ");
        }

        // (Background) Cell Color set by Conditional Formatting
        WorkbookEvaluatorProvider workbookEvaluatorProvider = (WorkbookEvaluatorProvider) cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        ConditionalFormattingEvaluator conditionalFormattingEvaluator = new ConditionalFormattingEvaluator(cell.getSheet().getWorkbook(), workbookEvaluatorProvider);
        List<EvaluationConditionalFormatRule> matchingCFRules = conditionalFormattingEvaluator.getConditionalFormattingForCell(cell);
        for (EvaluationConditionalFormatRule evalCFRule : matchingCFRules) {
            ConditionalFormattingRule cFRule = evalCFRule.getRule();
            if (cFRule.getPatternFormatting() != null) {
                XSSFColor backColor = (XSSFColor) cFRule.getPatternFormatting().getFillBackgroundColorColor();
                fileContent.append(" ").append("#").append(backColor.getARGBHex()).append(" ");
            } else if (cFRule.getColorScaleFormatting() != null) {
                XSSFColor[] colors = (XSSFColor[]) cFRule.getColorScaleFormatting().getColors();
                for (XSSFColor color : colors) {
                    fileContent.append(" ").append("#").append(color.getARGBHex()).append(" ");
                }
            }
        }
    }

    private void appendCellComment(Cell cell, StringBuilder fileContent) {
        String separator = "Comment:";
        String flags = "Flags: ";
        if (cell.getCellComment() != null) {
            String cellComment = cell.getCellComment().getString().getString();
            int beginSubString = cellComment.indexOf(separator);
            String substring;
            if (cellComment.contains(separator)) {
                if (cellComment.contains(flags)) {
                    substring = cellComment.substring(beginSubString, cellComment.indexOf(flags)).replace("\n", "");
                } else {
                    substring = cellComment.substring(beginSubString).replace("\n", "");
                }
            } else {
                substring = cellComment.substring(0, cellComment.indexOf('\n'));
            }
            fileContent.append("{").append(substring).append("} ");
        }
    }

    private List<String> getAllExcelFiles() throws IOException {
        List<String> listOfExcelFiles;

        Path ectRootPath = Paths.get(searchDirectory);
        if (!ectRootPath.toFile().isDirectory()) {
            throw new IllegalArgumentException("Path must be a directory !");
        }

        try (Stream<Path> walk = Files.walk(ectRootPath)) {
            listOfExcelFiles = walk
                    .filter(p -> !p.toFile().isDirectory())
                    .map(Path::toString)
                    .filter(containsTarget.negate().and(isExcelFile))
                    .collect(Collectors.toList());
        }
        return listOfExcelFiles;
    }

    private String printFileEmptyOrDamagedMessage(String pathToExcel) {
        StringBuilder warning = new StringBuilder(Paths.get(pathToExcel).toFile().toString());

        if (Paths.get(pathToExcel).toFile().length() == 0) {
            warning.append(" can't be opened. This excel file may be empty !");
        } else {
            warning.append(" can't be opened. This excel file may be damaged !");
        }
        return warning.toString();
    }

	public Log getLog() {
		return log;
	}
	
	private void removeEmptyColumnsOrRows(Sheet sheet) {
        List<List<String>> sheetTableLocal2 = new LinkedList<>();
        String sheetName = sheet.getSheetName();
        List<List<String>> sheetTableLocal = sheetTable.get(sheetName);
        for(List<String> list: sheetTableLocal) {
        	List<String> listNew = list.stream().filter(x->!x.trim().isEmpty()).collect(Collectors.toList());
        	if(listNew.size()>0) {
       			getLog().debug("listNew="+listNew);
        		sheetTableLocal2.add(list);
        	}
        }
        sheetTable.put(sheet.getSheetName(), sheetTableLocal2);
	}
	
	private Integer[] getMaxLength(List<String>[] arrList, String[][] arr) {
		Integer[] maxLengths = new Integer[arrList.length];
		for (int i = 0; i < arrList.length; i++) {
			String[] strArr = (String[]) arrList[i].toArray(new String[] {});
			arr[i] = strArr;
			Integer maxLength = arr[i].length;

			if (maxLengths[i] == null) {
				maxLengths[i] = maxLength;
			} else {
				maxLengths[i] = Math.max(maxLengths[i], maxLength);
			}
		}
		return maxLengths;
	}
	
	public String[][] createNewArray(String[][] arr, Integer[] maxLengths) {
		String[][] arrNew = null;
		for (int i = 0; i < arr.length; i++) {
			arrNew = new String[maxLengths[i]][];
			for (int j = 0; j < arr[i].length; j++) {
				arrNew[j] = new String[arr.length];
			}
		}
		return arrNew;
	}
	
	private void initNewArray(String[][] arrNew, String[][] arr) {
		for (int i = 0; i < arrNew.length; i++) {
			for (int j = 0; j < arr.length; j++) {
				if (arr[j].length > i) {
					String value = arr[j][i];
					arrNew[i][j] = value;
				} else {
					arrNew[i][j] = "";
				}
			}
		}
	}

	private void swapArrays(Sheet sheet) {
		String sheetName = sheet.getSheetName();
		List<List<String>> sheetTableLocal = sheetTable.get(sheetName);
		List<List<String>> sheetTableLocal2 = new LinkedList<>();

		String[][] arr = new String[sheetTableLocal.size()][];
		@SuppressWarnings("unchecked")
		List<String>[] arrList = sheetTableLocal.toArray(new List[] {});

		Integer[] maxLengths = getMaxLength(arrList, arr);

		String[][] arrNew = createNewArray(arr, maxLengths);
		
		initNewArray(arrNew, arr);
		
		for (int i = 0; i < arrNew.length; i++) {
			List<String> list = Arrays.asList(arrNew[i]);

			getLog().debug(list.toString());
			sheetTableLocal2.add(list);
		}
//      getLog().debug("sheetTableLocal2.size()="+sheetTableLocal2.size());
		sheetTable.put(sheet.getSheetName(), sheetTableLocal2);
	}
}
