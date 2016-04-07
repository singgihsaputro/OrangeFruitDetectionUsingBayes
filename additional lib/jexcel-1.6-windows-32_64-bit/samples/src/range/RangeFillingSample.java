/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package range;

import com.jniwrapper.win32.jexcel.Worksheet;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;

import javax.swing.*;
import java.awt.*;
import java.text.DateFormat;
import java.util.Date;

/**
 * This sample demonstrates bulk operations based on setValue2 Range interface.
 * Such approach provides a way for fast ranges filling.
 *
 * @author Igor Novikov
 */
public class RangeFillingSample
{
    public static void main(String[] args) throws Exception
    {
        //Simple JWorkbook-based application
        JFrame frame = new JFrame("Test application");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        Container cp = frame.getContentPane();
        cp.setLayout(new BorderLayout());
        JWorkbook workbook = new JWorkbook();
        cp.add(workbook);
        frame.setBounds(100, 100, 500, 500);
        frame.setVisible(true);

        Worksheet sheet = workbook.getActiveWorksheet();

        //Here we create 2d String array and fill according range
        final int ARRAY_SIZE = 10;
        int count = 0;
        final String[][] string_array = new String[ARRAY_SIZE][ARRAY_SIZE];
        for (int i = 0; i < ARRAY_SIZE; i++)
        {
            for (int j = 0; j < ARRAY_SIZE; j++)
            {
                string_array[i][j] = "result: " + Integer.toString(count);
                count++;
            }
        }
        sheet.fillWithArray("A1:J10", string_array);

        //Here we create 2d double array and fill according range
        double dcount = 0;
        final double[][] double_array = new double[ARRAY_SIZE][ARRAY_SIZE];
        for (int i = 0; i < ARRAY_SIZE; i++)
        {
            for (int j = 0; j < ARRAY_SIZE; j++)
            {
                double_array[i][j] = dcount++;
            }
        }
        sheet.fillWithArray("A12:J21", double_array);

        //The same for 2d long array
        long lcount = 0;
        final long[][] long_array = new long[ARRAY_SIZE][ARRAY_SIZE];
        for (int i = 0; i < ARRAY_SIZE; i++)
        {
            for (int j = 0; j < ARRAY_SIZE; j++)
            {
                long_array[i][j] = lcount++;
            }
        }
        sheet.fillWithArray("A23:J32", long_array);

        //and last sample for 2d Date array
        int date_count = 0;
        final Date[][] date_array = new Date[ARRAY_SIZE][ARRAY_SIZE];
        for (int i = 0; i < ARRAY_SIZE; i++)
        {
            for (int j = 0; j < ARRAY_SIZE; j++)
            {
                date_array[i][j] = new Date(System.currentTimeMillis() * date_count / 100);
                date_count++;
            }
        }
        sheet.fillWithArray("A34:J43", date_array, DateFormat.getDateInstance(DateFormat.MEDIUM));

    }

}