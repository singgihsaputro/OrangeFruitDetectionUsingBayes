/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package range;

import com.jniwrapper.win32.jexcel.Worksheet;
import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.GenericWorkbook;

/**
 * This sample demonstrates working with FormulaArray.
 */
public class FormulaArraySample
{
    public static void main(String[] args) throws Exception
    {
        Application application = new Application();
        application.setVisible(true);

        GenericWorkbook workbook = application.createWorkbook(null);
        Worksheet worksheet = workbook.getWorksheet(1);

        long[][] data = {{2, 4}, {5, 1}, {3, 4}, {6, 2}};
        worksheet.fillWithArray("A1:B4", data);
        com.jniwrapper.win32.jexcel.Range range = worksheet.getRange("C1:C4");
        range.setFormulaArray("=A1:A4+B1:B4");

        String formulaArray = range.getFormulaArray();
        System.out.println("formulaArray = " + formulaArray);

        System.out.println("Press 'Enter' to terminate applicaiton.");
        System.in.read();

        boolean forceQuit = true;
        application.setVisible(false);
        application.close(forceQuit);
    }
}
