/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package range;

import com.jniwrapper.win32.jexcel.*;

import java.util.List;

/**
 * @author Vladimir Kondrashchenko
 */
public class ReferencingRangesSample
{
    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);
        Worksheet worksheet = workbook.getWorksheet(1);

        //Referencing the range "A1:G12
        Range simpleRange = worksheet.getRange("A1:G12");

        //Not recommended
        Range compoundRange = worksheet.getRange("B1:B4;D11;H1:H13");

        //Creating compound ranges
        {
            Range range = worksheet.getRange("B1:B4");
            range.include("D11");
            range.include("H1:H13");
        }
        //More convenient way
        {
            Range range = worksheet.getRange("B1:B4").include("D11").include("H1:H13");
        }

        //Converting a cell to a range
        Cell cell = worksheet.getCell("A1");
        Range rangeFromCell = new Range(cell);

        //Converting a range to cells
        Range range = worksheet.getRange("B12:D12");
        List cells = range.getCells();
        for (int i = 0; i < cells.size(); i++)
        {
            Cell cellFromRange = (Cell) cells.get(i);
            System.out.println(cellFromRange);
        }

        application.close();

    }
}