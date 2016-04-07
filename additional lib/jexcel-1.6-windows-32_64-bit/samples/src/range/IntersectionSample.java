/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package range;

import com.jniwrapper.win32.jexcel.*;

/**
 * This sample demonstrates Range class ability to obtain ranges intersection,
 * cells merging and unmerging.
 *
 * @author Vladimir Kondrashchenko
 */
public class IntersectionSample
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

        return worksheet;
    }

    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);
        Worksheet worksheet = fillSampleData(workbook.getWorksheet(1));

        //Getting necessary ranges
        Range range1 = worksheet.getRange("A1:B3");
        Range range2 = worksheet.getRange("A2:C5");
        Range range3 = worksheet.getRange("D1:D4").include("A1");

        //Checking intersections
        if (range1.intersects(range2))
        {
            Range commonRange = range1.getIntersection(range2);
            System.out.println(range1 + " intersects " + range2 +
                    ". The common range is " + commonRange);
        }
        if (range1.intersects(range3))
        {
            System.out.println(range1 + " intersects " + range3 +
                    ". Commont range is " + range1.getIntersection(range3));
        }

        //Merging cells of range1
        range1.merge();

        //Unmerging cells of range1
        range1.unmerge();

        application.close();
    }
}