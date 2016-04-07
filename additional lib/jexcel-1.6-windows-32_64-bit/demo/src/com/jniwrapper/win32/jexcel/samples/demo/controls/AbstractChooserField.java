/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package com.jniwrapper.win32.jexcel.samples.demo.controls;

import javax.swing.JButton;
import javax.swing.JPanel;
import javax.swing.JTextField;
import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
 * @author Serge Piletsky
 */
public class AbstractChooserField extends JPanel implements ActionListener
{
    private JTextField _textField = new JTextField();
    private JButton _selectButton = new JButton("...");

    public AbstractChooserField()
    {
        Dimension preferredSize = new Dimension(150, 20);
        _textField.setPreferredSize(preferredSize);
        _textField.setMinimumSize(preferredSize);
        _selectButton.setPreferredSize(new Dimension(20, 20));

        setLayout(new BorderLayout());
        add(_textField, BorderLayout.CENTER);
        add(_selectButton, BorderLayout.EAST);
        _selectButton.addActionListener(this);
    }

    public JTextField getTextField()
    {
        return _textField;
    }

    public JButton getSelectButton()
    {
        return _selectButton;
    }

    public void setEnabled(boolean enabled)
    {
        _textField.setEnabled(enabled);
        _selectButton.setEnabled(enabled);
    }

    public void actionPerformed(ActionEvent e)
    {
    }
}