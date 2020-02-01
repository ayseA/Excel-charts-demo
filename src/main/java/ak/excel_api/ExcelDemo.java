package ak.excel_api;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTMarkerStyle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;

public class ExcelDemo {

	public static void makeChart(String outFile, String[] headers, double[]... dataStream) throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("chartDemo");
        
        int minL = Integer.MAX_VALUE;
        for (double[] d:dataStream)
        	minL = Math.min(minL, d.length);
        	
        if (minL>25) minL = 25;   // don't show too many-- demo only

        int nRows = minL;
        int nCols = dataStream.length;

        // column headers
        Row row = sheet.createRow(0);  
        for (int colInd = 1; colInd <= nCols; colInd++)
            row.createCell(colInd)
            	.setCellValue(headers[colInd-1]);

        for (int rowInd = 1; rowInd <= nRows; rowInd++) {
            row = sheet.createRow(rowInd);
            row.createCell(0).setCellValue(rowInd);
            for (int colInd = 1; colInd <= nCols; colInd++)
                row.createCell(colInd)
                	.setCellValue(dataStream[colInd-1][rowInd]);
        }

        Drawing drawing = sheet.createDrawingPatriarch();
        Chart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, nCols+2, 4, nCols+15, 22));
        chart.getOrCreateLegend().setPosition(LegendPosition.TOP_RIGHT);

        LineChartData data = chart.getChartDataFactory().createLineChartData();

        ChartAxis xAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        ValueAxis yAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        ChartDataSource<Number> cds, xs = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, nRows, 0, 0));
        for (int i=1; i<=nCols; i++) {
        	cds = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, nRows, i, i));
        	data.addSeries(xs, cds).setTitle(headers[i-1]);;
        }
        chart.plot(data, xAxis, yAxis);

        CTPlotArea plotArea = ((XSSFChart) chart).getCTChart().getPlotArea();
        CTMarker marker = CTMarker.Factory.newInstance();
        marker.setSymbol(CTMarkerStyle.Factory.newInstance());
        for (CTLineSer ser : plotArea.getLineChartArray()[0].getSerArray()) 
            ser.setMarker(marker);

        FileOutputStream out = new FileOutputStream(outFile);
        wb.write(out);
        wb.close();
        out.close();
    }
      
	@SuppressWarnings("deprecation")
	public static double[] readFile(String filePath, String columnHdr) {
		Sheet sheet = null;
		try (Workbook wb = new XSSFWorkbook(new FileInputStream(new File(filePath)))) {
			sheet = wb.getSheetAt(0);
		} catch (IOException e) {
			// TODO log
			e.printStackTrace();
		}
		
		Row headers = sheet.getRow(0);
		Cell theColmn = null;
		for (Cell c:headers)
			if (c.getStringCellValue().trim().equalsIgnoreCase(columnHdr))
				theColmn = c;		
		if (theColmn==null) 
			throw new IllegalArgumentException("Column "+columnHdr+" not found in "+filePath);		
		int colIndex = theColmn.getColumnIndex();
		if (!sheet.getRow(1).getCell(colIndex).getCellTypeEnum()
				.equals(CellType.NUMERIC)) 
			throw new IllegalArgumentException("Column "+columnHdr+" not numeric");
		
		double[] columnEntries  = new double[sheet.getPhysicalNumberOfRows()];
		int i = -1;
		for (Row row : sheet)
			if (i++>-1)
				columnEntries[i]=row.getCell(colIndex).getNumericCellValue();	
		return columnEntries;
	}

	@SuppressWarnings("deprecation")
	public static double[] readFileZ(String filePath, String columnHdr) throws FileNotFoundException, IOException {
		Workbook wb= new XSSFWorkbook(new FileInputStream(new File(filePath)));
		Sheet sheet = wb.getSheetAt(0);
		wb.close();
		Row headers = sheet.getRow(0);
		Cell theColmn = null;
		for (Cell c:headers)
			if (c.getStringCellValue().trim().equalsIgnoreCase(columnHdr))
				theColmn = c;

		if (theColmn==null) return null;
		final int colIndex = theColmn.getColumnIndex();
		if (!sheet.getRow(1).getCell(colIndex).getCellTypeEnum()
				.equals(CellType.NUMERIC)) 
			return null;
		
		double[] columnEntries  = new double[sheet.getPhysicalNumberOfRows()];
		int i = -1;
		for (Row row : sheet)
			if (i++>-1)  // parallel array-row indices
				columnEntries  [i]=row.getCell(colIndex).getNumericCellValue();	
		return columnEntries;
	}

	public static void main(String[] args) throws Exception {
		// Excel file sample link - https://go.microsoft.com/fwlink/?LinkID=521962
		String filePath = "Financial Sample.xlsx";  
		String[] columns = {"sales", "cogs", "profit"};
		double [] c1 = readFile(filePath, columns[0]);
		double [] c2 = readFile(filePath, columns[1]);
		double [] c3 = readFile(filePath, columns[2]);

		String out = "lineChart_"+System.currentTimeMillis()+".xlsx";
		makeChart(out, columns, c1, c2, c3);
	}
	
}
