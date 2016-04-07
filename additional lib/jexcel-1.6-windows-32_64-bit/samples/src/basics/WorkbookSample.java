/*
 * Copyright (c) 2000-2012 TeamDev Ltd. All rights reserved.
 * Use is subject to license terms.
 */

package basics;

import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.FileFormat;
import com.jniwrapper.win32.jexcel.GenericWorkbook;
import com.jniwrapper.win32.jexcel.Workbook;

import java.io.File;

/**
 * This sample shows how to read/modify workbook attributes, how to save workbook in Excel 2003 format,
 * and how to reopen workbook.
 *
 * The sample works with MS Excel in non-embedded mode.
 */
public class WorkbookSample
{
    public static void main(String[] args) throws Exception
    {
        //Start MS Excel application, crate workbook and make it visible.
        // Application starts invisible and without any workbooks
        Application application = new Application();
        Workbook workbook = application.createWorkbook("Custom title");

        printWorkbookAttributes(workbook);

        modifyWorkbookAttributes(workbook);

        File newFile = new File("Workbook.xls");
        //Save workbook in Excel 2003, to save in Excel 2007 format use FileFormat.OPENXMLWORKBOOK
        // format specificator and *.xlsx extension
        workbook.saveAs(newFile, FileFormat.WORKBOOKNORMAL, true);

        File workbookCopy = new File("WorkbookCopy.xls");
        workbook.saveCopyAs(workbookCopy);

        //Close workbook saving changes
        workbook.close(true);

        //Reopening the workbook
        workbook = application.openWorkbook(newFile, true, "xxx001");

        printWorkbookAttributes(workbook);

        workbook.close(false);

        //Perform cleanup after yourself and close the MS Excel application forcing it to quit
        application.close(true);
    }

    /**
     * Prints workbook attributes to console
     * @param workbook - workbook to print information about
     */
    public static void printWorkbookAttributes(GenericWorkbook workbook)
    {
        String fileName = workbook.getFile().getAbsolutePath();
        String name = workbook.getWorkbookName();
        String title = workbook.getTitle();
        String author = workbook.getAuthor();

        System.out.println("\n[Workbook Information]");
        System.out.println("File path: " + fileName);
        System.out.println("Name: " + name);
        System.out.println("Title: " + title);
        System.out.println("Author: " + author);

        if (workbook.hasPassword())
        {
            System.out.println("The workbook is protected with a password");
        }
        else
        {
            System.out.println("The workbook is not protected with a password");
        }
        if (workbook.isReadOnly())
        {
            System.out.println("Read only mode");
        }
    }

    /**
     * Modify workbook title, author and set password
     * @param workbook - workbook to modify attributes
     */
    public static void modifyWorkbookAttributes(GenericWorkbook workbook)
    {
        workbook.setTitle("X-files");
        workbook.setPassword("xxx001");
        workbook.setAuthor("Agent Smith");
    }
}