/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package cell;

import com.jniwrapper.win32.jexcel.*;

/**
 * This sample demonstraites how to work with named cells.
 *
 * @author Vladimir Kondrashchenko
 */
public class NamedCellsSample
{
    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);
        Worksheet worksheet = workbook.getWorksheet(1);

        //Obtaining a cell by its address
        Cell cell = worksheet.getCell("A1");

        //Printing cell name
        printCellName(cell);

        //Changing cell name
        cell.setName("New_Name");

        //Printing new cell name
        printCellName(cell);

        //Accessing a cell by its name
        cell = worksheet.getCell("New_Name");
        cell.setValue("String value");
        String value = cell.getString();
        System.out.println("\"" + cell.getName() + "\" cell value: " + value);

        application.close();
    }

    public static void printCellName(Cell cell)
    {
        if (cell.getName() == null)
        {
            System.out.println(cell.getAddress() + " name is not set up.");
        }
        else
        {
            System.out.println(cell.getAddress() + " name is \"" + cell.getName() + "\"");
        }
    }
}