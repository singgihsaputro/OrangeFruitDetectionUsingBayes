package nativepeer;

import com.jniwrapper.DoubleFloat;
import com.jniwrapper.Int32;
import com.jniwrapper.win32.automation.IDispatch;
import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.com.types.LocaleID;
import com.jniwrapper.win32.excel.*;
import com.jniwrapper.win32.excel.impl.ChartObjectsImpl;
import com.jniwrapper.win32.excel.impl.PivotFieldImpl;
import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.Workbook;
import com.jniwrapper.win32.jexcel.Worksheet;

/**
 * <p>  This sample demonstrates how to create a generic pivot table. It demonstrates how to use MS Excel API in cases when
 * <br> JExcel API does not cover some MS Excel functionality.
 * <br>
 * <br> This sample requires jexcel-full.jar in classpath
 *
 * @see <a href="http://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.workbook.pivottablewizard%28v=vs.80%29.aspx">
 *                   MSDN: Workbook.PivotTableWizard Method </a>
 */
public class PivotTableSample
{
    //File name where to save the result. Change to the one specific for your environment.
    private static final String FILE_NAME = "F:\\pivot.xlsx";

    public static void main(String[] args) throws Exception
    {
        //Start MS Excel
        Application excelApp = new Application();
        excelApp.setVisible(true);

        //Create test workbook
        Workbook workbook = excelApp.createWorkbook("Pivot Table Test");

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

        //<p>   Create the worksheet that will be used as the destination for pivot table.
        // <br>Place it right after the first one.
        final Worksheet worksheet2 = workbook.addWorksheet(worksheet1,"PivotTableSheet");

        //To avoid marshaling problems call to the native peer methods must be executed through the OleMessageLoop
        excelApp.getOleMessageLoop().doInvokeAndWait(new Runnable()
        {
            public void run()
            {
                //Create an unspecified variant value to fill-in those parameters which you want to skip.
                Variant unspecified = Variant.createUnspecifiedParameter();

                //Create variants with true and false values, since they will  be reused
                Variant variantFalse = new Variant(false);
                Variant variantTrue = new Variant(true);

                //Get native peers of Range and Cell. This is com.jniwrapper.win32.excel.Range interface
                Range sourceDataNativePeer = worksheet1.getRange("A1:C4").getPeer();
                Range destinationNativePeer = worksheet2.getCell("A1").getPeer();

                //Get the Worksheet native peer, since the pivotTableWizard is a MS Excel API method, but not the JExcel API method.
                _Worksheet worksheetNativePeer = worksheet1.getPeer();

                //Call the pivot table wizard method.
                PivotTable table = worksheetNativePeer.pivotTableWizard(new Variant(XlPivotTableSourceType.xlDatabase),  //SourceType
                                                        new Variant(sourceDataNativePeer), //SourceData
                                                        new Variant(destinationNativePeer), //TableDestination
                                                        new Variant("PivotTable1"), //TableName
                                                        variantFalse, //RowGrand
                                                        variantFalse, //ColumnGrand
                                                        variantTrue, //SaveData
                                                        variantTrue, //HasAutoFormat
                                                        unspecified, //AutoPage
                                                        unspecified, //Reserved
                                                        variantFalse, //BackgroundQuery
                                                        variantFalse, //OptimizeCache
                                                        new Variant(XlOrder.xlDownThenOver), //PageFieldOrder
                                                        unspecified, //PageFieldWrapCount
                                                        unspecified, //ReadData
                                                        unspecified, //Connection
                                                        new Int32(LocaleID.LOCALE_SYSTEM_DEFAULT));

                IDispatch dispatch = worksheetNativePeer.chartObjects(unspecified, new Int32(LocaleID.LOCALE_SYSTEM_DEFAULT));

                ChartObjectsImpl chartObjects = new ChartObjectsImpl(dispatch);
                ChartObject chartObject = chartObjects.add(new DoubleFloat(100), new DoubleFloat(150), new DoubleFloat(300), new DoubleFloat(225));

                _Chart chart = chartObject.getChart();
                chart.setSourceData(sourceDataNativePeer, new Variant(XlRowCol.xlRows));

                chart.setAutoDelete(false);
                chart.release();

                chartObject.setAutoDelete(false);
                chartObject.release();

                chartObjects.setAutoDelete(false);
                chartObjects.release();

                dispatch.setAutoDelete(false);
                dispatch.release();

                //Now get the data fields which were added at pivot table creation as hidden and set their representation

                //The Customer field will be the second in list of fields but we want to place it as the first one.
                //Get field's dispatch interface
                IDispatch fieldCustomerDispatch = table.getHiddenFields(new Variant(2));
                //Then query for the PivotTable interface
                PivotFieldImpl fieldCustomer = new PivotFieldImpl(fieldCustomerDispatch);
                //Set this field as the RowLabels
                fieldCustomer.setOrientation(new XlPivotFieldOrientation(XlPivotFieldOrientation.xlRowField));
                //Place it as the first displayable field
                fieldCustomer.setPosition(new Variant(1));

                //Now do the same for Data and Sales, considering that displayable fields are removed from the list of hidden fields.

                IDispatch fieldDataDispatch = table.getHiddenFields(new Variant(1));
                PivotFieldImpl fieldData = new PivotFieldImpl(fieldDataDispatch);
                fieldData.setOrientation(new XlPivotFieldOrientation(XlPivotFieldOrientation.xlRowField));
                fieldData.setPosition(new Variant(2));

                IDispatch fieldSalesDispatch = table.getHiddenFields(new Variant(1));
                PivotFieldImpl fieldSales = new PivotFieldImpl(fieldSalesDispatch);
                //Set sales as data field, so it is used as 'Sum of Values' field
                fieldSales.setOrientation(new XlPivotFieldOrientation(XlPivotFieldOrientation.xlDataField));
            }
        });

        //Save workbook
//        workbook.saveAs(new File(FILE_NAME), FileFormat.WORKBOOKDEFAULT, true);

        System.in.read();

        //Close the MS Excel application.
        workbook.close(false);
        excelApp.close();
    }

}
