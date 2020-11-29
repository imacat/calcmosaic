/* Copyright (c) 2012 imacat
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/* 
 * OptionDialog
 * 
 * Created on 2012-10-29
 * 
 * Copyright (c) 2012 imacat
 */

package tw.idv.imacat.calcmosaic;

import java.io.File;
import java.util.ResourceBundle;

import java.awt.Dialog.ModalityType;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.GridLayout;
import java.awt.Toolkit;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JSpinner;
import javax.swing.SpinnerNumberModel;
import javax.swing.UIManager;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 * An dialog to obtain the options.
 * 
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 2.1.0
 */
public class OptionDialog {
    
    /** The localization resources. */
    private ResourceBundle l10n = null;
    
    /** The option set. */
    private Options options = null;
    
    /** The dialog. */
    private JDialog dialog = new JDialog(
        null, ModalityType.TOOLKIT_MODAL);
    
    /** The image files. */
    public File files[] = new File[0];;
    
    /** The number of rows to use. */
    public int numRows = 0;
    
    /** Whether we should run in slow motion (for presentation). */
    public boolean isSlow = false;
    
    /** The current directory. */
    private File currentDirectory = null;
    
    /** The file chooser button. */
    private JButton filesButton = new JButton();
    
    /** The spinner of the number of rows. */
    private JSpinner numRowsSpinner = new JSpinner();
    
    /** The check box of whether we should run in slow motion. */
    private JCheckBox isSlowCheckBox = new JCheckBox();
    
    /** The OK button. */
    private JButton okButton = new JButton();
    
    /**
     * Creates a new option dialog.
     * 
     * @param files   the image files
     * @param numRows the number of rows to use
     * @param isSlow  whether we should run in slow motion (for
     *                presentation)
     */
    public OptionDialog(File files[], int numRows, boolean isSlow) {
        // Obtains the localization resources
        this.l10n = ResourceBundle.getBundle(
            this.getClass().getPackage().getName() + ".res.L10n");
        this.files = files;
        this.numRows = numRows;
        this.isSlow = isSlow;
    }
    
    /**
     * Asks the user and returns the options.
     * 
     * @return the options set by the user
     */
    public Options getOptions() {
        JPanel optionPanel = new JPanel();
        JLabel label = null;
        JButton cancelButton = new JButton(_("dialog_btn_cancel"));
        SpinnerNumberModel spinnerModel = new SpinnerNumberModel(
            this.numRows, 0, 32767, 1);
        Dimension screenSize
            = Toolkit.getDefaultToolkit().getScreenSize();
        Dimension frameSize = null;
        
        // Makes "enter" acts as "space" that activates the buttons
        UIManager.put("Button.defaultButtonFollowsFocus", true);
        
        this.dialog.setIconImage(CalcMosaic.icon);
        this.dialog.setTitle(String.format(_("dialog_opts_title"),
            CalcMosaic.APP_NAME));
        this.dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        optionPanel.setLayout(new GridLayout(4, 2, 10, 5));
        
        // Reserves the space for a 3-digit number
        this.filesButton.setText(
            String.format(_("dialog_btn_files"), 256));
        this.filesButton.setMnemonic(KeyEvent.VK_F);
        this.filesButton.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent e) {
                    chooseFiles();
                }
            });
        label = new JLabel(_("dialog_label_files"));
        label.setLabelFor(this.filesButton);
        optionPanel.add(label);
        optionPanel.add(this.filesButton);
        
        this.numRowsSpinner.setModel(spinnerModel);
        label = new JLabel(_("dialog_label_numrows"));
        label.setLabelFor(this.numRowsSpinner);
        optionPanel.add(label);
        optionPanel.add(this.numRowsSpinner);
        
        this.isSlowCheckBox.setText(_("dialog_checkbox_isslow"));
        this.isSlowCheckBox.setSelected(this.isSlow);
        label = new JLabel(_("dialog_label_isslow"));
        label.setLabelFor(this.isSlowCheckBox);
        optionPanel.add(label);
        optionPanel.add(this.isSlowCheckBox);
        
        this.okButton.setText(_("dialog_btn_ok"));
        this.okButton.setMnemonic(KeyEvent.VK_O);
        this.okButton.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent e) {
                    options = new Options(files,
                        (Integer) numRowsSpinner.getValue(),
                        isSlowCheckBox.isSelected());
                    dialog.dispose();
                }
            });
        if (this.files.length == 0) {
            this.okButton.setEnabled(false);
        } else {
            this.okButton.setEnabled(true);
        }
        optionPanel.add(this.okButton);
        cancelButton = new JButton(_("dialog_btn_cancel"));
        cancelButton.setMnemonic(KeyEvent.VK_C);
        cancelButton.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent e) {
                    options = null;
                    dialog.dispose();
                }
            });
        optionPanel.add(cancelButton);
        
        this.options = null;
        this.dialog.add(optionPanel);
        this.dialog.pack();
        frameSize = this.dialog.getSize();
        this.dialog.setLocation(
                (screenSize.width - frameSize.width) / 2,
                (screenSize.height - frameSize.height) / 2
            );
        this.filesButton.setText(
            String.format(_("dialog_btn_files"), this.files.length));
        this.dialog.setVisible(true);
        
        return this.options;
    }
    
    /**
     * Chooses the image files.
     * 
     */
    private void chooseFiles() {
        File chosen[] = new File[0];
        
        while (true) {
            JFileChooser chooser = new JFileChooser(
                this.currentDirectory);
            FileNameExtensionFilter filter = new FileNameExtensionFilter(
                _("dialog_file_chooser_image_type"), "jpg", "png", "gif");
            int result = -1;
            
            chooser.setDialogTitle(_("dialog_file_chooser_title"));
            chooser.setFileFilter(filter);
            chooser.setMultiSelectionEnabled(true);
            result = chooser.showOpenDialog(this.dialog);
            this.currentDirectory = chooser.getCurrentDirectory();
            
            if (result == JFileChooser.CANCEL_OPTION) {
                return;
            
            } else if (result == JFileChooser.ERROR_OPTION) {
                JOptionPane.showMessageDialog(this.dialog,
                    _("err_file_chooser_failed"));
                return;
            }
            
            chosen = chooser.getSelectedFiles();
            if (chosen.length > 256) {
                JOptionPane.showMessageDialog(this.dialog,
                    _("err_file_too_many"),
                    String.format(_("dialog_err_title"),
                        CalcMosaic.APP_NAME),
                    JOptionPane.ERROR_MESSAGE);
                continue;
            }
            break;
        }
        
        this.files = chosen;
        this.filesButton.setText(
            String.format(_("dialog_btn_files"), this.files.length));
        if (this.files.length == 0) {
            this.okButton.setEnabled(false);
        } else {
            this.okButton.setEnabled(true);
            this.okButton.grabFocus();
        }
        return;
    }
    
    /**
     * Gets a string for the given key from this resource bundle
     * or one of its parents.  If the key is missing, returns
     * the key itself.
     * 
     * @param key the key for the desired string 
     * @return the string for the given key 
     */
    private String _(String key) {
        try {
            return new String(
                this.l10n.getString(key).getBytes("ISO-8859-1"),
                "UTF-8");
        } catch (java.io.UnsupportedEncodingException e) {
            return this.l10n.getString(key);
        } catch (java.util.MissingResourceException e) {
            return key;
        }
    }
}
