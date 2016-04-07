/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package basics;

import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.ExcelException;
import com.jniwrapper.win32.jexcel.GenericWorkbook;
import com.jniwrapper.win32.jexcel.Worksheet;

import java.util.List;

/**
 * This sample demonstrates how to obtain a worksheet by its index or name,
 * add, move or remove a worksheet.
 *
 * @author Vladimir Kondrashchenko
 */
public class WorksheetsSample
{
    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();

        GenericWorkbook workbook = application.createWorkbook(null);

        Worksheet sheet2 = workbook.getWorksheet("Sheet2");
        workbook.addWorksheet(sheet2, "Custom sheet");

        printWorksheets(workbook);

        //Obtaining a worksheet by its name
        Worksheet customSheet = workbook.getWorksheet("Custom sheet");
        //Obtaining a worksheet by its index
        int lastIndex = workbook.getWorksheetCount();
        Worksheet lastWorksheet = workbook.getWorksheet(lastIndex);

        workbook.moveWorksheet(customSheet, lastWorksheet);

        if (customSheet.getIndex() == workbook.getWorksheetCount())
        {
            System.out.println(customSheet.getName() + " is the last worksheet.");
        }

        workbook.moveWorksheet(customSheet, null);

        if (customSheet.getIndex() == 1)
        {
            System.out.println(customSheet.getName() + " is the first worksheet.");
        }

        workbook.removeWorksheet(customSheet);

        printWorksheets(workbook);

        application.close();

    }

    public static void printWorksheets(GenericWorkbook workbook)
    {
        List worksheets = workbook.getWorksheets();

        System.out.println("List of worksheets: ");
        for (int i = 0; i < worksheets.size(); i++)
        {
            Worksheet worksheet = (Worksheet) worksheets.get(i);
            System.out.println(worksheet.getName());
        }
    }
}