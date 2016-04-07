/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package range;

import com.jniwrapper.win32.jexcel.*;

import java.util.Date;
import java.util.GregorianCalendar;

/**
 * This sample demonstrates how to find values of different types.
 *
 * @author Vladimir Kondrashchenko
 */
public class FindValuesSample
{
    private static Worksheet fillSampleData(Worksheet worksheet)
    {
        worksheet.getCell(1, 1).setValue("Apple");
        worksheet.getCell(2, 1).setValue("Grapefruit");
        worksheet.getCell(3, 1).setValue("Strawberry");
        worksheet.getCell(4, 1).setValue("Grapefruit");
        worksheet.getCell(5, 1).setValue("Apple");

        worksheet.getCell(1, 2).setValue(12);
        worksheet.getCell(2, 2).setValue(15);
        worksheet.getCell(3, 2).setValue(10);
        worksheet.getCell(4, 2).setValue(2);
        worksheet.getCell(5, 2).setValue("=SUM(B1:B4)");

        worksheet.getCell(1, 3).setValue(1.1);
        worksheet.getCell(2, 3).setValue(0.23);
        worksheet.getCell(3, 3).setValue(5.1);
        worksheet.getCell(4, 3).setValue(2);
        worksheet.getCell(5, 3).setValue("=SUM(C1:C4)");

        GregorianCalendar calendar = new GregorianCalendar(2000, 0, 12);
        worksheet.getCell(1, 4).setValue(calendar.getTime());
        worksheet.getCell(1, 4).setColumnWidth(25);
        return worksheet;
    }

    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);
        Worksheet worksheet = fillSampleData(workbook.getWorksheet(1));

        //Getting necessary range
        Range range = worksheet.getRange("A1:C5").include("D1");

        //Specifying find attributes
        Range.SearchAttributes searchAttributes = new Range.SearchAttributes();
        searchAttributes.setCaseSensitive(true);
        searchAttributes.setLookIn(Range.FindLookIn.VALUES);
        searchAttributes.setForwardDirection(true);

        //Looking for a string value
        String strValue = "Grapefruit";
        Cell cell = range.find(strValue, searchAttributes);
        if (cell == null)
        {
            System.out.println("\"" + strValue + "\" was not found.");
        }
        else
        {
            System.out.println("\"" + strValue + "\" was found in " + cell.getAddress());
        }
        cell = range.findNext();
        if (cell == null)
        {
            System.out.println("\"" + strValue + "\" was not found.");
        }
        else
        {
            System.out.println("\"" + strValue + "\" was found in " + cell.getAddress());
        }

        strValue = "Tomato";
        cell = range.find(strValue, searchAttributes);
        if (cell == null)
        {
            System.out.println("\"" + strValue + "\" was not found.");
        }
        else
        {
            System.out.println("\"" + strValue + "\" was found in " + cell.getAddress());
        }

        //Looking for a Date value
        Date dateValue = new GregorianCalendar(2000, 0, 12).getTime();
        cell = range.find(dateValue, searchAttributes);
        if (cell == null)
        {
            System.out.println(dateValue + " was not found.");
        }
        else
        {
            System.out.println(dateValue + " was found in " + cell.getAddress());
        }

        //Looking for a long calculated value
        long longValue = 39;
        cell = range.find(longValue, searchAttributes);
        if (cell == null)
        {
            System.out.println(longValue + " was not found.");
        }
        else
        {
            System.out.println(longValue + " was found in " + cell.getAddress());
        }

        application.close();
    }
}