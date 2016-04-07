package com.jniwrapper.win32.jexcel.samples.demo;

import com.jniwrapper.win32.com.ComException;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.io.FileNotFoundException;
import java.io.File;

/**
 *
 */
class RecentFilesAction extends AbstractAction
{

    private final JExcelDemo _demo;
    private final JWorkbook _workbook;
    private final File _file;

    RecentFilesAction(String name, JExcelDemo demo, JWorkbook workbook, File file)
    {
        super(name);
        _demo = demo;
        _workbook = workbook;
        _file = file;
    }

    RecentFilesAction(JExcelDemo demo, JWorkbook workbook, File file)
    {
        this._demo = demo;
        this._workbook = workbook;
        this._file = file;
    }

    public void actionPerformed(ActionEvent e)
    {
        String message = "";
        if (getDemo().getExcelVersion() < JExcelDemo.VERSION_NUMBER_2007)
        {
            message = "Excel version installed on this os can open only .xls files";
        }

        try
        {
            checkPreviewMode();
            getWorkbook().openWorkbook(getFile());
            if (getWorkbook().getFile() != null)
            {
                getDemo().setTitle(JExcelDemo.DEMO_TITLE + " - [" + getWorkbook().getFile().getAbsolutePath() + "]");
            }

        }
        catch (FileNotFoundException ex)
        {
            JOptionPane.showMessageDialog(SwingUtilities.windowForComponent(getDemo()),
                    "Cannot find file \"" + getFile().getAbsolutePath() + "\".", "File not found",
                    JOptionPane.ERROR_MESSAGE);
        }
        catch (ComException exp)
        {
            JOptionPane.showMessageDialog(SwingUtilities.windowForComponent(getDemo()),
                    "Cannot open file \"" +getFile().getAbsolutePath() + "\n" + message,
                    "Unsupported format",
                    JOptionPane.ERROR_MESSAGE);
        }
        catch (IllegalArgumentException ex)
        {
            JOptionPane.showMessageDialog(SwingUtilities.windowForComponent(getDemo()),
                    "Cannot open file \"" + getFile().getAbsolutePath() + "\n" + message,
                    "Unsupported format",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private boolean checkPreviewMode()
    {
        if (getWorkbook().isPrintPreview())
        {
            JOptionPane.showMessageDialog(SwingUtilities.windowForComponent(getDemo()),
                    "Unable to process the command while the preview mode is active.",
                    "JExcel Demo",
                    JOptionPane.ERROR_MESSAGE);
            return true;
        }
        else
            return false;        
    }

    private JExcelDemo getDemo()
    {
        return _demo;
    }

    private JWorkbook getWorkbook()
    {
        return _workbook;
    }

    private File getFile()
    {
        return _file;
    }
}
