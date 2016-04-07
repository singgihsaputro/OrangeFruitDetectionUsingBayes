package nativepeer;

import com.jniwrapper.DoubleFloat;
import com.jniwrapper.Int32;
import com.jniwrapper.win32.automation.IDispatch;
import com.jniwrapper.win32.automation.types.BStr;
import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.com.types.LocaleID;
import com.jniwrapper.win32.excel.*;
import com.jniwrapper.win32.excel.impl.ChartObjectsImpl;
import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.Workbook;
import com.jniwrapper.win32.jexcel.Worksheet;

/**
 * <p>  This sample demonstrates how to export an Excel chart to image. It demonstrates how to use MS Excel API in cases when
 * <br> JExcel API does not cover some MS Excel functionality.
 * <br>
 * <br> This sample requires jexcel-full.jar in classpath
 */
public class ExportChartToImage
{
    public static void main(String[] args) throws Exception
    {
        //Start Excel
        Application excelApp = new Application();
        excelApp.setVisible(true);

        //Create test workbook
        Workbook workbook = excelApp.createWorkbook("Chart Test");

        //Get the first (and the only) worksheet
        final Worksheet worksheet1 = workbook.getWorksheet(1);

        //Fill-in the first worksheet with sample data
        worksheet1.getCell("A1").setValue("Date");
        worksheet1.getCell("A2").setValue("March 1");
        worksheet1.getCell("A3").setValue("March 8");
        worksheet1.getCell("A4").setValue("March 15");

        worksheet1.getCell("B1").setValue("Customer");
        worksheet1.getCell("B2").setValue("Smith");
        worksheet1.getCell("B3").setValue("Jones");
        worksheet1.getCell("B4").setValue("James");

        worksheet1.getCell("C1").setValue("Sales");
        worksheet1.getCell("C2").setValue("23");
        worksheet1.getCell("C3").setValue("17");
        worksheet1.getCell("C4").setValue("39");

        excelApp.getOleMessageLoop().doInvokeAndWait(new Runnable()
        {
            public void run()
            {
                final Variant unspecified = Variant.createUnspecifiedParameter();
                final Int32 localeID = new Int32(LocaleID.LOCALE_SYSTEM_DEFAULT);

                Range sourceDataNativePeer = worksheet1.getRange("A1:C4").getPeer();
                _Worksheet worksheetNativePeer = worksheet1.getPeer();

                IDispatch chartObjectDispatch = worksheetNativePeer.chartObjects(unspecified, localeID);

                ChartObjectsImpl chartObjects = new ChartObjectsImpl(chartObjectDispatch);
                ChartObject chartObject = chartObjects.add(new DoubleFloat(100), new DoubleFloat(150), new DoubleFloat(300), new DoubleFloat(225));

                _Chart chart = chartObject.getChart();
                chart.setSourceData(sourceDataNativePeer, new Variant(XlRowCol.xlRows));

                BStr fileName = new BStr("C:\\Temp\\chart.gif");
                Variant filterName = new Variant("gif");
                Variant interactive = new Variant(false);

                chart.export(fileName, filterName, interactive);

                chart.setAutoDelete(false);
                chart.release();

                chartObject.setAutoDelete(false);
                chartObject.release();

                chartObjects.setAutoDelete(false);
                chartObjects.release();

                chartObjectDispatch.setAutoDelete(false);
                chartObjectDispatch.release();
            }
        });

        System.out.println("Press 'Enter' to terminate the application");
        System.in.read();

        //Close the MS Excel application.
        boolean saveChanges = false;
        workbook.close(saveChanges);
        boolean forceQuit = true;
        excelApp.close(forceQuit);
    }

}
