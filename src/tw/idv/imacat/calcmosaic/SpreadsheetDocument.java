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
 * SpreadsheetDocument
 * 
 * Created on 2012-08-16, rewritten from CalcMosaic
 * 
 * Copyright (c) 2012 imacat
 */

package tw.idv.imacat.calcmosaic;

import java.awt.Dimension;
import java.util.ResourceBundle;

import com.sun.star.awt.Size;
import com.sun.star.awt.XView;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.script.XLibraryContainer;
import com.sun.star.script.provider.XScript;
import com.sun.star.script.provider.XScriptProvider;
import com.sun.star.script.provider.XScriptProviderSupplier;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.XColumnRowRange;
import com.sun.star.table.XTableRows;
import com.sun.star.uno.UnoRuntime;

// Used in slow-motion mode.
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.awt.XControlModel;
import com.sun.star.container.XIndexContainer;
import com.sun.star.drawing.XControlShape;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPageSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.form.XFormComponent;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.script.ScriptEventDescriptor;
import com.sun.star.script.XEventAttacherManager;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.sheet.XSpreadsheetView;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XColumnRowRange;
import com.sun.star.table.XTableColumns;
import com.sun.star.table.XTableRows;

/**
 * A spreadsheet document.
 * 
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 2.1.0
 */
public class SpreadsheetDocument {
    
    /** The BASIC library name. */
    private static final String libraryName = "Standard";
    
    /** The localization resources. */
    private ResourceBundle l10n = null;
    
    /** The spreadsheet document. */
    private XSpreadsheetDocument document = null;
    
    /** The spreadsheet collection. */
    private XSpreadsheets sheets = null;
    
    /** The document controller. */
    private XController xController = null;
    
    /** The spreadsheet view. */
    private XSpreadsheetView xSpreadsheetView = null;
    
    /** The row height, to be cached. */
    private int rowHeight = -1;
    
    /**
     * Creates a new instance of SpreadsheetDocument.
     * 
     * @param document the document component
     */
    public SpreadsheetDocument(XComponent document) {
        XModel xModel = null;
        
        // Obtains the localization resources
        this.l10n = ResourceBundle.getBundle(
            this.getClass().getPackage().getName() + ".res.L10n");
        
        this.document = (XSpreadsheetDocument)
            UnoRuntime.queryInterface(
            XSpreadsheetDocument.class, document);
        
        this.sheets = this.document.getSheets();
        
        // Obtains the document controller
        xModel = (XModel) UnoRuntime.queryInterface(
            XModel.class, this.document);
        this.xController = xModel.getCurrentController();
        this.xSpreadsheetView = (XSpreadsheetView)
            UnoRuntime.queryInterface(
            XSpreadsheetView.class, this.xController);
    }
    
    /**
     * Adds a module of BASIC macros to the document.
     * 
     * @param module the name of the module.
     * @param macros the content of the module
     */
    public void addBasicModule(String module, String macros) {
        XPropertySet xProps = null;
        Object libraries = null;
        XLibraryContainer libraryContainer = null;
        Object library = null;
        XNameContainer moduleContainer = null;
        
        xProps = (XPropertySet) UnoRuntime.queryInterface(
            XPropertySet.class, this.document);
        try {
            libraries = xProps.getPropertyValue("BasicLibraries");
        } catch (com.sun.star.beans.UnknownPropertyException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
        libraryContainer = (XLibraryContainer)
            UnoRuntime.queryInterface(
            XLibraryContainer.class, libraries);
        if (!libraryContainer.hasByName(libraryName)) {
            try {
                libraryContainer.createLibrary(libraryName);
            } catch (com.sun.star.lang.IllegalArgumentException e) {
                throw new java.lang.IllegalArgumentException(e);
            } catch (com.sun.star.container.ElementExistException e) {
                throw new java.lang.RuntimeException(e);
            }
        }
        try {
            library = libraryContainer.getByName(libraryName);
        } catch (com.sun.star.container.NoSuchElementException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
        moduleContainer = (XNameContainer) UnoRuntime.queryInterface(
            XNameContainer.class, library);
        if (moduleContainer.hasByName(module)) {
            try {
                moduleContainer.removeByName(module);
            } catch (com.sun.star.container.NoSuchElementException e) {
                throw new java.lang.RuntimeException(e);
            } catch (com.sun.star.lang.WrappedTargetException e) {
                throw new java.lang.RuntimeException(e);
            }
        }
        try {
            moduleContainer.insertByName(module, macros);
        } catch (com.sun.star.lang.IllegalArgumentException e) {
            throw new java.lang.IllegalArgumentException(e);
        } catch (com.sun.star.container.ElementExistException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
    }
    
    /**
     * Runs a BASIC macro of the document.
     * 
     * @param module the name of the module of the macro.
     * @param macro the name of the macro.
     * @param param the parameters for the macro
     */
    public void runBasicMacro(String module, String macro, Object param[]) {
        String scriptUri = String.format(
            "vnd.sun.star.script:%s.%s.%s?language=Basic&location=document",
            libraryName, module, macro);
        XScriptProviderSupplier xSupplier = (XScriptProviderSupplier)
            UnoRuntime.queryInterface(
            XScriptProviderSupplier.class, document);
        XScriptProvider xProvider = xSupplier.getScriptProvider();
        XScript xScript = null;
        short outParamIndex[][] = new short[0][0];
        Object outParam[][] = new Object[0][0];
        
        try {
            xScript = xProvider.getScript(scriptUri);
        } catch (com.sun.star.script.provider.ScriptFrameworkErrorException e) {
            throw new java.lang.IllegalArgumentException(e);
        }
        try {
            xScript.invoke(param, outParamIndex, outParam);
        } catch (com.sun.star.script.provider.ScriptFrameworkErrorException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.reflection.InvocationTargetException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.DisposedException e) {
        }
    }
    
    /**
     * Returns the size of the window.
     * 
     * @return the size of the window
     */
    public Dimension getWindowSize() {
        XFrame frame = this.xController.getFrame();
        XWindow window = frame.getComponentWindow();
        XView xView = (XView) UnoRuntime.queryInterface(
            XView.class, window);
        Size size = xView.getSize();
        return new Dimension(size.Width - 55, size.Height - 72);
    }
    
    /**
     * Returns the height of the rows.  We do not need
     * the width, since the height is always shorter in new
     * spreadsheets.
     * 
     * @return the height of the rows.
     */
    public int getRowHeight() {
        if (this.rowHeight != -1) {
            return this.rowHeight;
        }
        
        XIndexAccess sheetIndexAccess = (XIndexAccess)
            UnoRuntime.queryInterface(
            XIndexAccess.class, this.sheets);
        Object sheet = null;
        XColumnRowRange xColumnRowRange = null;
        XTableRows rows = null;
        XPropertySet xProps = null;
        
        try {
            sheet = sheetIndexAccess.getByIndex(0);
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new java.lang.IndexOutOfBoundsException(
                e.getMessage());
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
        
        // Obtains the height of the rows.
        xColumnRowRange = (XColumnRowRange) UnoRuntime.queryInterface(
            XColumnRowRange.class, sheet);
        rows = xColumnRowRange.getRows();
        xProps = (XPropertySet) UnoRuntime.queryInterface(
            XPropertySet.class, rows);
        try {
            this.rowHeight = (Integer)
                xProps.getPropertyValue("Height");
        } catch (com.sun.star.beans.UnknownPropertyException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
        
        return this.rowHeight;
    }
    
    /**
     * Creates a mosaic spreadsheet.
     * 
     * @param name   the spreadsheet name
     * @param index  the spreadsheet index
     * @param colors the color values matrix
     */
    public void createMosaicSheet(String name, int index,
            int colors[][]) {
        XSpreadsheet sheet = this.initializeMosaicSheet(name, index);
        
        this.xSpreadsheetView.setActiveSheet(sheet);
        this.setColumnWidth(sheet, colors.length, colors[0].length);
        
        // Paints the colors.
        for (int row = 0; row < colors.length; row++) {
            for (int column = 0; column < colors[row].length; column++) {
                XCell cell = null;
                XPropertySet xProps = null;
                
                try {
                    cell = sheet.getCellByPosition(column, row);
                } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
                    throw new java.lang.IndexOutOfBoundsException(
                        e.getMessage());
                }
                xProps = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, cell);
                try {
                    xProps.setPropertyValue(
                        "CellBackColor", colors[row][column]);
                    xProps.setPropertyValue(
                        "IsCellBackgroundTransparent", false);
                } catch (com.sun.star.beans.UnknownPropertyException e) {
                    throw new java.lang.RuntimeException(e);
                } catch (com.sun.star.beans.PropertyVetoException e) {
                    throw new java.lang.RuntimeException(e);
                } catch (com.sun.star.lang.IllegalArgumentException e) {
                    throw new java.lang.IllegalArgumentException(e);
                } catch (com.sun.star.lang.WrappedTargetException e) {
                    throw new java.lang.RuntimeException(e);
                }
            }
        }
        return;
    }
    
    /**
     * Initializes a mosaic spreadsheet.
     * 
     * @param name  the spreadsheet name
     * @param index the spreadsheet index
     * @return the spreadsheet
     */
    private XSpreadsheet initializeMosaicSheet(
            String name, int index) {
        XIndexAccess sheetIndexAccess = (XIndexAccess)
            UnoRuntime.queryInterface(
            XIndexAccess.class, this.sheets);
        Object sheet = null;
        XSpreadsheet xSheet = null;
        
        // Adds our spreadsheet.
        if (this.sheets.hasByName(name)) {
            if (sheetIndexAccess.getCount() == 1) {
                this.sheets.insertNewByName(name + 1, (short) 0);
                try {
                    this.sheets.removeByName(name);
                } catch (com.sun.star.container.NoSuchElementException e) {
                    throw new java.lang.RuntimeException(e);
                } catch (com.sun.star.lang.WrappedTargetException e) {
                    throw new java.lang.RuntimeException(e);
                }
                this.sheets.insertNewByName(name, (short) index);
                try {
                    this.sheets.removeByName(name + 1);
                } catch (com.sun.star.container.NoSuchElementException e) {
                    throw new java.lang.RuntimeException(e);
                } catch (com.sun.star.lang.WrappedTargetException e) {
                    throw new java.lang.RuntimeException(e);
                }
            } else {
                try {
                    this.sheets.removeByName(name);
                } catch (com.sun.star.container.NoSuchElementException e) {
                    throw new java.lang.RuntimeException(e);
                } catch (com.sun.star.lang.WrappedTargetException e) {
                    throw new java.lang.RuntimeException(e);
                }
                this.sheets.insertNewByName(name, (short) index);
            }
        } else {
            this.sheets.insertNewByName(name, (short) index);
        }
        // Cleans-up
        if (index == 0) {
            String names[] = this.sheets.getElementNames();
            for (int i = 0; i < names.length; i++) {
                if (!names[i].equals(name)) {
                    try {
                        this.sheets.removeByName(names[i]);
                    } catch (com.sun.star.container.NoSuchElementException e) {
                        throw new java.lang.RuntimeException(e);
                    } catch (com.sun.star.lang.WrappedTargetException e) {
                        throw new java.lang.RuntimeException(e);
                    }
                }
            }
        }
        
        // Obtains the speadsheet.
        try {
            sheet = sheets.getByName(name);
        } catch (com.sun.star.container.NoSuchElementException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
        xSheet = (XSpreadsheet) UnoRuntime.queryInterface(
            XSpreadsheet.class, sheet);
        return xSheet;
    }
    
    /**
     * Sets the width of the columns.
     * 
     * @param sheet   the spreadsheet
     * @param rows    the number of rows in the cell range
     * @param columns the number of columns in the cell range
     */
    private void setColumnWidth(XSpreadsheet sheet,
            int rows, int columns) {
        XCellRange range = null;
        XColumnRowRange xColumnRowRange = null;
        XTableColumns xColumns = null;
        int height = -1;
        XPropertySet xProps = null;
        
        try {
            range = sheet.getCellRangeByPosition(
                0, 0, columns - 1, rows - 1);
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new java.lang.IndexOutOfBoundsException(
                e.getMessage());
        }
        xColumnRowRange = (XColumnRowRange) UnoRuntime.queryInterface(
            XColumnRowRange.class, range);
        xColumns = xColumnRowRange.getColumns();
        xProps = (XPropertySet) UnoRuntime.queryInterface(
            XPropertySet.class, xColumns);
        try {
            xProps.setPropertyValue("Width", this.getRowHeight());
        } catch (com.sun.star.beans.UnknownPropertyException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.beans.PropertyVetoException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.IllegalArgumentException e) {
            throw new java.lang.IllegalArgumentException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
    }
    
    /**
     * Adds the macro triggers to show the sheets.
     * 
     */
    public void addMacroTriggers() {
        XIndexAccess sheetIndexAccess = (XIndexAccess)
            UnoRuntime.queryInterface(
            XIndexAccess.class, this.sheets);
        int count = sheetIndexAccess.getCount();
        Object firstSheet = null, lastSheet = null;
        XSpreadsheet xFirstSheet = null, xLastSheet;
        XCell breaksCell = null, musicCell = null;
        XPropertySet breaksCellProps = null, musicCellProps = null;
        int breaksColor = -1, musicColor = -1;
        
        // Adds start buttons on the first and the last sheets.
        try {
            firstSheet = sheetIndexAccess.getByIndex(0);
            lastSheet = sheetIndexAccess.getByIndex(count - 1);
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new java.lang.IndexOutOfBoundsException(
                e.getMessage());
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
        xFirstSheet = (XSpreadsheet) UnoRuntime.queryInterface(
            XSpreadsheet.class, firstSheet);
        xLastSheet = (XSpreadsheet) UnoRuntime.queryInterface(
            XSpreadsheet.class, lastSheet);
        this.addStartButton(xFirstSheet);
        this.addStartButton(xLastSheet);
        
        // Sets the text color to be the same as the background color
        // at B1 and C1 of the first sheet, where are the breaks
        // between sheets and the music files, so that users can
        // hide these two pieces of information.
        this.xSpreadsheetView.setActiveSheet(xFirstSheet);
        try {
            breaksCell = xFirstSheet.getCellByPosition(1, 0);
            musicCell = xFirstSheet.getCellByPosition(2, 0);
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new java.lang.IndexOutOfBoundsException(
                e.getMessage());
        }
        breaksCellProps = (XPropertySet) UnoRuntime.queryInterface(
            XPropertySet.class, breaksCell);
        musicCellProps = (XPropertySet) UnoRuntime.queryInterface(
            XPropertySet.class, musicCell);
        try {
            breaksColor = (Integer) breaksCellProps.getPropertyValue(
                "CellBackColor");
            musicColor = (Integer) musicCellProps.getPropertyValue(
                "CellBackColor");
        } catch (com.sun.star.beans.UnknownPropertyException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
        try {
            breaksCellProps.setPropertyValue(
                "CharColor", breaksColor);
            musicCellProps.setPropertyValue(
                "CharColor", musicColor);
        } catch (com.sun.star.beans.UnknownPropertyException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.beans.PropertyVetoException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.IllegalArgumentException e) {
            throw new java.lang.IllegalArgumentException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
    }
    
    /**
     * Adds a start button to show the sheets.
     * 
     * @param sheet the spreadsheet to add a start button unto
     */
    private void addStartButton(XSpreadsheet sheet) {
        XMultiServiceFactory serviceManager = (XMultiServiceFactory)
            UnoRuntime.queryInterface(
            XMultiServiceFactory.class, this.document);
        Object shape = null;
        XControlShape xShape = null;
        Point position = new Point();
        Size size = new Size();
        Object button = null;
        XCell cell = null;
        XPropertySet xProps = null;
        int color = -1;
        XControlModel buttonModel = null;
        XDrawPageSupplier xDrawPageSupplier = null;
        XDrawPage drawPage = null;
        ScriptEventDescriptor event = new ScriptEventDescriptor();
        XFormComponent xFormComponent = null;
        Object form = null;
        XIndexContainer xContainer = null;
        XEventAttacherManager xEventAttacherManager = null;
        
        // Creates the shape for the button, as a control shape.
        try {
            shape = serviceManager.createInstance(
                "com.sun.star.drawing.ControlShape");
        } catch (com.sun.star.uno.Exception e) {
            throw new java.lang.RuntimeException(e);
        }
        xShape = (XControlShape) UnoRuntime.queryInterface(
            XControlShape.class, shape);
        position.X = 0;
        position.Y = 0;
        xShape.setPosition(position);
        size.Width = this.getRowHeight();
        size.Height = size.Width;
        try {
            xShape.setSize(size);
        } catch (com.sun.star.beans.PropertyVetoException e) {
            throw new java.lang.RuntimeException(e);
        }
        
        // Creates the button.
        try {
            button = serviceManager.createInstance(
                "com.sun.star.form.component.CommandButton");
        } catch (com.sun.star.uno.Exception e) {
            throw new java.lang.RuntimeException(e);
        }
        try {
            cell = sheet.getCellByPosition(0, 0);
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            throw new java.lang.IndexOutOfBoundsException(
                e.getMessage());
        }
        xProps = (XPropertySet) UnoRuntime.queryInterface(
            XPropertySet.class, cell);
        try {
            color = (Integer) xProps.getPropertyValue(
                "CellBackColor");
        } catch (com.sun.star.beans.UnknownPropertyException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
        xProps = (XPropertySet) UnoRuntime.queryInterface(
            XPropertySet.class, button);
        try {
            xProps.setPropertyValue("BackgroundColor", color);
        } catch (com.sun.star.beans.UnknownPropertyException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.beans.PropertyVetoException e) {
            throw new java.lang.RuntimeException(e);
        } catch (com.sun.star.lang.IllegalArgumentException e) {
            throw new java.lang.IllegalArgumentException(e);
        } catch (com.sun.star.lang.WrappedTargetException e) {
            throw new java.lang.RuntimeException(e);
        }
        
        // Adds the button to the page.
        buttonModel = (XControlModel) UnoRuntime.queryInterface(
            XControlModel.class, button);
        xShape.setControl(buttonModel);
        xDrawPageSupplier = (XDrawPageSupplier) UnoRuntime.queryInterface(
            XDrawPageSupplier.class, sheet);
        drawPage = xDrawPageSupplier.getDrawPage();
        drawPage.add(xShape);
        
        // Sets the action of the button
        event.ListenerType = "XActionListener";
        event.EventMethod = "actionPerformed";
        event.ScriptType = "Script";
        event.ScriptCode = "vnd.sun.star.script:"
            + "Standard._0Play.subPlaySheets"
            + "?language=Basic&location=document";
        xFormComponent = (XFormComponent) UnoRuntime.queryInterface(
            XFormComponent.class, button);
        form = xFormComponent.getParent();
        xContainer = (XIndexContainer) UnoRuntime.queryInterface(
            XIndexContainer.class, form);
        xEventAttacherManager = (XEventAttacherManager)
            UnoRuntime.queryInterface(
            XEventAttacherManager.class, form);
        try {
            xEventAttacherManager.registerScriptEvent(
                xContainer.getCount() - 1, event);
        } catch (com.sun.star.lang.IllegalArgumentException e) {
            throw new java.lang.IllegalArgumentException(e);
        }
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
