/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package range;

import com.jniwrapper.win32.jexcel.*;

import java.util.GregorianCalendar;

/**
 * This sample demonstrates how to sort ranges.
 *
 * @author Vladimir Kondrashchenko
 */
public class SortValuesSample
{
    private static Worksheet fillSampleData(Worksheet worksheet)
    {
        worksheet.getCell(1, 1).setValue("Apple");
        worksheet.getCell(2, 1).setValue("Orange");
        worksheet.getCell(3, 1).setValue("Strawberry");
        worksheet.getCell(4, 1).setValue("Grapefruit");
        worksheet.getCell(5, 1).setValue("Apple");

        worksheet.getCell(1, 2).setValue(12);
        worksheet.getCell(2, 2).setValue(15);
        worksheet.getCell(3, 2).setValue(10);
        worksheet.getCell(4, 2).setValue(2);
        worksheet.getCell(5, 2).setValue(100);

        worksheet.getCell(1, 3).setValue(1.1);
        worksheet.getCell(2, 3).setValue(0.23);
        worksheet.getCell(3, 3).setValue(5.1);
        worksheet.getCell(4, 3).setValue(2);
        worksheet.getCell(5, 3).setValue(0.01);

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
        Range range = worksheet.getRange("A1:D5");

        //Prinitng range values before sorting
        System.out.println("Range Values Before Sorting:");
        for (int i = 1; i <= 5; i++)
        {
            for (int j = 1; j <= 4; j++)
            {
                String value = worksheet.getCell(i, j).getString();
                if (value != null)
                {
                    System.out.print('\t' + value);
                }
            }
            System.out.print('\n');
        }

        //Sorting the range by column "A"
        range.sort("A", true, true);

        System.out.println("\nRange Values After Sorting by Column \"A\":");
        for (int i = 1; i <= 5; i++)
        {
            for (int j = 1; j <= 4; j++)
            {
                String value = worksheet.getCell(i, j).getString();
                if (value != null)
                {
                    System.out.print('\t' + value);
                }
            }
            System.out.print('\n');
        }

        application.close();
    }
}