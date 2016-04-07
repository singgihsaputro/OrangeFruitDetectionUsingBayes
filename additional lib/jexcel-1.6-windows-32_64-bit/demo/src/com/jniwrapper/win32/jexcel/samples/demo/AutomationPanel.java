/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package com.jniwrapper.win32.jexcel.samples.demo;

import com.jniwrapper.win32.jexcel.ui.JWorkbook;
import com.jniwrapper.win32.jexcel.Range;
import com.jniwrapper.win32.jexcel.format.TextAlignment;
import com.jniwrapper.win32.jexcel.format.TextOrientation;
import com.jniwrapper.win32.jexcel.samples.demo.controls.ChooseColorField;
import com.jniwrapper.win32.com.ComException;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.util.List;

/**
 * @author Vladimir Kondrashchenko
 */
class AutomationPanel extends JPanel
{

    private static final int RECENT_FILES_COUNT = 10;

    private static final TextOrientation[] TEXT_ORIENTATIONS =
            {TextOrientation.VERTICAL, TextOrientation.HORIZONTAL, TextOrientation.DOWNWARD, TextOrientation.UPWARD};
    private static final TextAlignment[] VERTICAL_ALIGNMENT =
            {TextAlignment.TOP, TextAlignment.CENTER, TextAlignment.BOTTOM, TextAlignment.JUSTIFY, TextAlignment.DISTRIBUTED};
    private static final TextAlignment[] HORIZONTAL_ALIGNMENT =
            {TextAlignment.GENERAL, TextAlignment.LEFT, TextAlignment.CENTER, TextAlignment.RIGHT, TextAlignment.FILL,
                    TextAlignment.JUSTIFY, TextAlignment.DISTRIBUTED};

    private JTextField taSelection;
    private ChooseColorField _cellsBGColor;
    private ChooseColorField _cellsFontColor;
    private JComboBox _textOrientation;
    private JComboBox _verticalAlignment;
    private JComboBox _horisontalAlignment;
    private JPanel _recentFilesPanel;
    private JTextField _documentTitle;
    private JTextField _author;
    private final JFrame _parent;

    private Range _selection;
    private final JWorkbook _workbook;

    AutomationPanel(JWorkbook workbook, JFrame parent)
    {
        super(null);
//        super(new GridBagLayout());
        _workbook = workbook;
        _parent = parent;

        JLabel lblSelectionCaption = new JLabel("Selection:");
        taSelection = new JTextField("");
        taSelection.setEnabled(false);

        JPanel cellFormat = createCellFormatPanel();

        _recentFilesPanel = new JPanel(new GridBagLayout());
        _recentFilesPanel.setBorder(BorderFactory.createTitledBorder(null, "Recently Opened Files",
                TitledBorder.DEFAULT_JUSTIFICATION, TitledBorder.DEFAULT_POSITION, null, Color.DARK_GRAY));
        _recentFilesPanel.setPreferredSize(new Dimension(300, 300));

        JPanel summaryPanel = createSummaryPanel();

//        add(summaryPanel, new GridBagConstraints(0, 0, 1, 1, 1.0, 0.0,
//                GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL, new Insets(10, 5, 0, 5), 0, 0));
//        add(_recentFilesPanel, new GridBagConstraints(0, 1, 1, 1, 1.0, 0.0,
//                GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL, new Insets(5, 5, 0, 5), 0, 0));
//        add(lblSelectionCaption, new GridBagConstraints(0, 2, 1, 1, 0.0, 0.0,
//                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5, 5, 0, 0), 0, 0));
//        add(taSelection, new GridBagConstraints(0, 3, 1, 1, 1.0, 0.0,
//                GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL, new Insets(0, 5, 0, 5), 0, 0));
//        add(cellFormat, new GridBagConstraints(0, 4, 1, 1, 1.0, 0.0,
//                GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL, new Insets(5, 5, 0, 5), 0, 0));
//        add(new JPanel(), new GridBagConstraints(0, 5, 1, 1, 1.0, 1.0,
//                GridBagConstraints.WEST, GridBagConstraints.BOTH, new Insets(0, 0, 0, 0), 0, 0));
        add(summaryPanel);
        add(_recentFilesPanel);
        add(lblSelectionCaption);
        add(taSelection);
        add(cellFormat);

        final int offset = 5;
        final int componentWidth = 230;
        summaryPanel.setBounds(offset, 0, componentWidth, 105);
        _recentFilesPanel.setBounds(offset, 110, componentWidth, 240);
        lblSelectionCaption.setBounds(offset, 355, componentWidth, 20);
        taSelection.setBounds(offset, 375, componentWidth, 20);
        cellFormat.setBounds(offset, 400, componentWidth, 190);
    }

    private JPanel createSummaryPanel()
    {
        JLabel lblTitle = new JLabel("Document Title:");
        _documentTitle = new JTextField();
        _documentTitle.setEnabled(false);
        JLabel lblAuthor = new JLabel("Author:");
        _author = new JTextField("Author:");
        _author.setEnabled(false);

        JPanel summaryPanel = new JPanel(new GridBagLayout());
        summaryPanel.setBorder(BorderFactory.createTitledBorder(null, "Workbook Summary",
                TitledBorder.DEFAULT_JUSTIFICATION, TitledBorder.DEFAULT_POSITION, null, Color.DARK_GRAY));
        summaryPanel.add(lblTitle, new GridBagConstraints(0, 0, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 5, 0, 5), 0, 0));
        summaryPanel.add(_documentTitle, new GridBagConstraints(0, 1, 1, 1, 1.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL, new Insets(0, 5, 0, 5), 0, 0));
        summaryPanel.add(lblAuthor, new GridBagConstraints(0, 2, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5, 5, 0, 5), 0, 0));
        summaryPanel.add(_author, new GridBagConstraints(0, 3, 1, 1, 1.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL, new Insets(0, 5, 5, 5), 0, 0));

        summaryPanel.setPreferredSize(new Dimension(300, 300));

        return summaryPanel;
    }

    private JPanel createCellFormatPanel()
    {
        JPanel cellFormat = new JPanel(new GridBagLayout());
        cellFormat.setBorder(BorderFactory.createTitledBorder(null, "Cell Format",
                TitledBorder.DEFAULT_JUSTIFICATION, TitledBorder.DEFAULT_POSITION, null, Color.DARK_GRAY));

        JLabel lblCellsBGColor = new JLabel("Background Color: ");
        JLabel lblCellsFontColor = new JLabel("Font Color: ");
        _cellsBGColor = new ChooseColorField();
        _cellsBGColor.addPropertyChangeListener(ChooseColorField.PROPERTY_COLOR, new PropertyChangeListener()
        {
            public void propertyChange(PropertyChangeEvent evt)
            {
                if (_selection != null)
                {
                    try
                    {
                        _selection.setInteriorColor(_cellsBGColor.getColor());
                    }
                    catch (ComException ex)
                    {
                        showErrorMessage();
                    }
                }
            }
        });
        _cellsFontColor = new ChooseColorField();
        _cellsFontColor.addPropertyChangeListener(ChooseColorField.PROPERTY_COLOR, new PropertyChangeListener()
        {
            public void propertyChange(PropertyChangeEvent evt)
            {
                if (_selection != null)
                {
                    com.jniwrapper.win32.jexcel.format.Font font = _selection.getFont();
                    font.setColor(_cellsFontColor.getColor());
                    try
                    {
                        _selection.setFont(font);
                    }
                    catch (ComException ex)
                    {
                        showErrorMessage();
                    }
                }
            }
        });
        JButton btnClear = new JButton(new AbstractAction("Clear Cells")
        {
            public void actionPerformed(ActionEvent e)
            {
                if (_selection != null)
                {
                    try
                    {
                        _selection.clear();
                    }
                    catch (ComException ex)
                    {
                        showErrorMessage();
                    }
                }
            }
        });
        JPanel clearPanel = new JPanel();
        clearPanel.add(btnClear);

        JLabel lblTextOrientation = new JLabel("Text Orientation:");
        _textOrientation = createTextOrientationField();

        JLabel lblVerticalAlignment = new JLabel("Vertical Alignment:");
        _verticalAlignment = createVerticalAlignmentField();

        JLabel lblHorizontalAlignment = new JLabel("Horizontal Alignment:");
        _horisontalAlignment = createHorizontalAlignmentField();

        cellFormat.add(lblCellsBGColor, new GridBagConstraints(0, 0, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 5, 0, 5), 0, 0));
        cellFormat.add(_cellsBGColor, new GridBagConstraints(1, 0, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 0, 0, 5), 0, 0));
        cellFormat.add(lblCellsFontColor, new GridBagConstraints(0, 1, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5, 5, 0, 5), 0, 0));
        cellFormat.add(_cellsFontColor, new GridBagConstraints(1, 1, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5, 0, 0, 5), 0, 0));
        cellFormat.add(lblTextOrientation, new GridBagConstraints(0, 2, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5, 5, 0, 5), 0, 0));
        cellFormat.add(_textOrientation, new GridBagConstraints(1, 2, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5, 0, 0, 5), 0, 0));
        cellFormat.add(lblVerticalAlignment, new GridBagConstraints(0, 3, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5, 5, 0, 5), 0, 0));
        cellFormat.add(_verticalAlignment, new GridBagConstraints(1, 3, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5, 0, 0, 5), 0, 0));
        cellFormat.add(lblHorizontalAlignment, new GridBagConstraints(0, 4, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5, 5, 0, 5), 0, 0));
        cellFormat.add(_horisontalAlignment, new GridBagConstraints(1, 4, 1, 1, 0.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5, 0, 0, 5), 0, 0));
        cellFormat.add(clearPanel, new GridBagConstraints(0, 5, 2, 1, 1.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL, new Insets(5, 5, 0, 5), 0, 0));

        return cellFormat;
    }

    private JComboBox createTextOrientationField()
    {
        JComboBox result = new JComboBox(TEXT_ORIENTATIONS);
        result.setPreferredSize(new Dimension(100, 20));
        result.addItemListener(new ItemListener()
        {
            public void itemStateChanged(ItemEvent e)
            {
                if (_selection == null || e.getStateChange() != ItemEvent.SELECTED)
                {
                    return;
                }
                try
                {
                    _selection.setTextOrientation((TextOrientation) e.getItem());
                }
                catch (ComException ex)
                {
                    showErrorMessage();
                }
            }
        });

        return result;
    }

    private JComboBox createVerticalAlignmentField()
    {
        JComboBox result = new JComboBox(VERTICAL_ALIGNMENT);
        result.setPreferredSize(new Dimension(100, 20));
        result.addItemListener(new ItemListener()
        {
            public void itemStateChanged(ItemEvent e)
            {
                if (_selection == null || e.getStateChange() != ItemEvent.SELECTED)
                {
                    return;
                }
                try
                {
                    _selection.setVerticalAlignment((TextAlignment) e.getItem());
                }
                catch (ComException ex)
                {
                    showErrorMessage();
                }
            }
        });

        return result;
    }

    private JComboBox createHorizontalAlignmentField()
    {
        JComboBox result = new JComboBox(HORIZONTAL_ALIGNMENT);
        result.setPreferredSize(new Dimension(100, 20));
        result.addItemListener(new ItemListener()
        {
            public void itemStateChanged(ItemEvent e)
            {
                if (_selection == null || e.getStateChange() != ItemEvent.SELECTED)
                {
                    return;
                }
                try
                {
                    _selection.setHorizontalAlignment((TextAlignment) e.getItem());
                }
                catch (ComException ex)
                {
                    showErrorMessage();
                }
            }
        });

        return result;
    }

    private void showErrorMessage()
    {
        SwingUtilities.invokeLater(new Runnable()
        {
            public void run()
            {
                JOptionPane.showMessageDialog(SwingUtilities.windowForComponent(AutomationPanel.this),
                        "Unable to process the command while a cell has the keyboard focus.",
                        "JExcel Demo",
                        JOptionPane.ERROR_MESSAGE);
            }
        });
    }

    void updateRecentFiles()
    {
        _recentFilesPanel.removeAll();

        List files = _workbook.getApplication().getRecentFiles();

        int recentFiles = Math.min(files.size(), RECENT_FILES_COUNT);

        for (int i = 0; i < recentFiles; i++)
        {
            final File file = (File) files.get(i);
            final int maxLength = 30;
            String pathText = JExcelDemo.pathText(file, maxLength);
            LinkLabel recentFile = new LinkLabel(pathText,
                    new ImageIcon(AutomationPanel.class.getResource("res/excel.gif")),
                    new RecentFilesAction((JExcelDemo) _parent, _workbook, file));            
            _recentFilesPanel.add(recentFile, new GridBagConstraints(0, i, 1, 1, 0.0, 0.0,
                    GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(2, 0, 2, 0), 0, 0));
        }
        _recentFilesPanel.add(new JPanel(), new GridBagConstraints(0, files.size(), 1, 1, 1.0, 0.0,
                GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL, new Insets(0, 0, 0, 0), 0, 0));

        _documentTitle.setText(_workbook.getTitle());
        _author.setText(_workbook.getAuthor());

        revalidate();
        repaint();
    }

    public void setSelection(Range range)
    {
        _selection = range;
        if (range == null)
        {
            taSelection.setText("");
            return;
        }
        taSelection.setText(_selection.getAddress());
        if (range.getInteriorColor() != null)
        {
            _cellsBGColor.setColor(range.getInteriorColor());
        }
        else
        {
            _cellsBGColor.setColor(Color.WHITE);
        }
        if (range.getFont().getColor() != null)
        {
            _cellsFontColor.setColor(range.getFont().getColor());
        }
        else
        {
            _cellsFontColor.setColor(Color.WHITE);
        }

        _textOrientation.setSelectedItem(_selection.getTextOrientation());
        _verticalAlignment.setSelectedItem(_selection.getVerticalAlignment());
        _horisontalAlignment.setSelectedItem(_selection.getHorizontalAlignment());

        SwingUtilities.invokeLater(new Runnable()
        {
            public void run()
            {
                repaint();
            }
        });
    }
}