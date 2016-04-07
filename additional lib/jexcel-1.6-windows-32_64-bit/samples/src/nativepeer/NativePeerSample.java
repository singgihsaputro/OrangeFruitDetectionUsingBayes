/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer;

import com.jniwrapper.win32.automation.types.BStr;
import com.jniwrapper.win32.excel._Application;
import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.ExcelException;

import java.lang.reflect.InvocationTargetException;

/**
 * This sample demonstrates how to work with Excel native peers. Native peers allow
 * to implement Excel features that are not present in the current version of JExcel.
 *
 * @author Vladimir Kondrashchenko
 */
public class NativePeerSample
{
    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();

        System.out.println("Current decimal separator: " + getDecimalSeparator(application));

        setDecimalSeparator(application, ",");

        System.out.println("New decimal separator: " + getDecimalSeparator(application));

        application.close();
    }

    /**
     * This function obtains current decimal separator using Application native peer.
     *
     * @param application is an instance of Excel {@link Application}.
     * @return current decimal separator.
     */
    public static String getDecimalSeparator(final Application application)
    {
        final String[] result = new String[1];
        Runnable runnable = new Runnable()
        {
            public void run()
            {
                _Application nativeApp = application.getPeer();
                BStr decSeparator = nativeApp.getDecimalSeparator();
                result[0] = decSeparator.getValue();
            }
        };
        try
        {
            application.getOleMessageLoop().doInvokeAndWait(runnable);
        }
        catch (InterruptedException e)
        {
            e.printStackTrace();
        }
        catch (InvocationTargetException e)
        {
            e.printStackTrace();
        }

        return result[0];
    }

    /**
     * This function changes decimal separator using Application native peer.
     *
     * @param application  is an instance of Excel {@link Application}.
     * @param newSeparator is a new separator.
     */
    public static void setDecimalSeparator(final Application application, final String newSeparator)
    {
        Runnable runnable = new Runnable()
        {
            public void run()
            {
                _Application nativeApp = application.getPeer();
                BStr decSeparator = new BStr(newSeparator);
                nativeApp.setDecimalSeparator(decSeparator);
            }
        };
        try
        {
            application.getOleMessageLoop().doInvokeAndWait(runnable);
        }
        catch (InterruptedException e)
        {
            e.printStackTrace();
        }
        catch (InvocationTargetException e)
        {
            e.printStackTrace();
        }
    }
}