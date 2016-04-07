/*
 * Copyright (c) 2000-2012 TeamDev Ltd. All rights reserved.
 * Use is subject to license terms.
 */

package basics;

import com.jniwrapper.win32.jexcel.*;

import java.io.File;
import java.io.IOException;

/**
 * This sample demonstrates how to copy worksheet from one workbook to another and
 * how to save workbook in MS Excel 2003 format.
 *
 * MS Excel workbook with at least three worksheets stored in MS Excel 2003 format with name test.xls
 * must be present in the project working directory.
 *
 * The sample works with MS Excel in non-embedded mode.
 *
 * @author Dmitriy Milashenko
 */
public class WorksheetCopySample
{
    public static void main(String[] args) throws ExcelException, IOException
    {
        //Start MS Excel application.
        // Application starts invisible and without any workbooks
        Application application = new Application();

        Workbook wb1 = application.openWorkbook(new File("D:\\tmp\\test.xls"));
        Workbook wb2 = application.createWorkbook("WB2");
        Workbook wb3 = application.createWorkbook("WB3");

        Worksheet ws2 = wb1.getWorksheet(2);
        Worksheet ws3 = wb1.getWorksheet(3);

        //Copy worksheets to workbooks and set them as the first worksheet.
        // For this pass the first workbook worksheet as the 'before' parameter
        wb2.copyWorksheet(ws2,wb2.getWorksheet(1),null);
        wb3.copyWorksheet(ws3,wb3.getWorksheet(1),null);

        //Save workbook in Excel 2003, to save in Excel 2007 format use FileFormat.OPENXMLWORKBOOK
        // format specificator and *.xlsx extention
        wb2.saveAs(new File("D:\\tmp\\test2.xls"), FileFormat.WORKBOOKNORMAL, true);
        wb3.saveAs(new File("D:\\tmp\\test3.xls"), FileFormat.WORKBOOKNORMAL, true);

        //Perform cleanup after yourself and close the MS Excel application forcing it to quit
        application.close(true);
    }
}