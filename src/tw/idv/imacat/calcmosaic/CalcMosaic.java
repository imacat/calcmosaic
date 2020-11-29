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
 * CalcMosaic
 * 
 * Created on 2012-08-12
 * 
 * Copyright (c) 2012 imacat
 */

package tw.idv.imacat.calcmosaic;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.awt.Toolkit;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;
import javax.imageio.ImageIO;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

/**
 * A program to create mosaic art from images with Calc.
 * 
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 2.1.0
 */
public class CalcMosaic {
    
    /** The title of the application. */
    public static final String APP_NAME = "Calc Moasic";
    
    /** The title of the application. */
    private static final String ICON = "images/calcmosaic.png";
    
    /** The version number. */
    private static final String VERSION = "2.1.0";
    
    /** The name of the jar file. */
    private static final String JAR_NAME = "calcmosaic.jar";
    
    /** The module name of the BASIC playing macros. */
    private static final String playModule = "_0Play";
    
    /** The module name of the BASIC macros. */
    private static final String basicModule = "_1CalcMosaic";
    
    /** The BASIC macros to play the mosaic art. */
    private static final String playMacros =
"' Copyright (c) 2012 imacat.\n"
+ "' \n"
+ "' Licensed under the Apache License, Version 2.0 (the \"License\");\n"
+ "' you may not use this file except in compliance with the License.\n"
+ "' You may obtain a copy of the License at\n"
+ "' \n"
+ "'     http://www.apache.org/licenses/LICENSE-2.0\n"
+ "' \n"
+ "' Unless required by applicable law or agreed to in writing, software\n"
+ "' distributed under the License is distributed on an \"AS IS\" BASIS,\n"
+ "' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.\n"
+ "' See the License for the specific language governing permissions and\n"
+ "' limitations under the License.\n"
+ "\n"
+ "' _0Play: The BASIC macros to play the spreadsheets\n"
+ "'   Created with " + APP_NAME + " version " + VERSION + "\n"
+ "'   By imacat <imacat@mail.imacat.idv.tw>, 2012/10/5\n"
+ "\n"
+ "Option Explicit\n"
+ "\n"
+ "' The default background music, related to the document file\n"
+ "Const sDefaultMusic = \"\"\n"
+ "' The default break time between sheets (ms)\n"
+ "Const nDefaultBreak = 100\n"
+ "\n"
+ "' subPlaySheets: Plays the spreadsheets.\n"
+ "'   1. Save file first for the macro to work.\n"
+ "'   2. Hit the start button at A1 of the first\n"
+ "'      and last sheet to play the sheets.\n"
+ "'   3. You can set the break time (in\n"
+ "'      milliseconds) between sheets at B1 of\n"
+ "'      the first sheet (default 100 ms).\n"
+ "'   4. You can set the background music at C1\n"
+ "'      of the first sheet.  This only works on\n"
+ "'      Linux and you must install mplayer.\n"
+ "'      Relative file path is relative to the\n"
+ "'      document file.  Play once before your\n"
+ "'      live show to avoid the I/O delay with cache.\n"
+ "Sub subPlaySheets\n"
+ "\tDim oController As Object, oSheets As Object\n"
+ "\tDim nBreak As Integer, nI As Integer\n"
+ "\tDim nCount As Integer, oSheet As Object\n"
+ "\t\n"
+ "\toController = ThisComponent.getCurrentController\n"
+ "\toSheets = ThisComponent.getSheets\n"
+ "\tnBreak = oSheets.getByIndex (0).getCellByPosition (1, 0).getValue\n"
+ "\tIf nBreak = 0 Then\n"
+ "\t\tnBreak = nDefaultBreak\n"
+ "\tEnd If\n"
+ "\tnCount = oSheets.getCount\n"
+ "\t\n"
+ "\tsubPlayMusic\n"
+ "\t\n"
+ "\tFor nI = 0 To nCount - 1\n"
+ "\t\toSheet = oSheets.getByIndex (nI)\n"
+ "\t\toController.setActiveSheet (oSheet)\n"
+ "\t\tWait nBreak\n"
+ "\tNext nI\n"
+ "End Sub\n"
+ "\n"
+ "' subPlayMusic: Plays the background music.\n"
+ "Sub subPlayMusic\n"
+ "\tDim oSheet As Object, sMusic As String\n"
+ "\tDim sFile As String, nPos As Integer\n"
+ "\t\n"
+ "\t' This only works on Linux and you must install mplayer.\n"
+ "\tIf Not FileExists (\"/usr/bin/mplayer\") Then\n"
+ "\t\tExit Sub\n"
+ "\tEnd if\n"
+ "\t\n"
+ "\toSheet = ThisComponent.getSheets.getByIndex (0)\n"
+ "\tsMusic = oSheet.getCellByPosition (2, 0).getString\n"
+ "\tIf sMusic = \"\" Then\n"
+ "\t\tsMusic = sDefaultMusic\n"
+ "\tEnd If\n"
+ "\tIf sMusic = \"\" Then\n"
+ "\t\tExit Sub\n"
+ "\tEnd If\n"
+ "\t\n"
+ "\t' Relative file path is relative to the document file.\n"
+ "\tIf Left (sMusic, 1) <> \"/\" Then\n"
+ "\t\tsFile = ConvertFromURL (ThisComponent.getLocation)\n"
+ "\t\tIf sFile = \"\" Then\n"
+ "\t\t\tExit Sub\n"
+ "\t\tEnd If\n"
+ "\t\tnPos = Len (sFile)\n"
+ "\t\tDo While  Mid (sFile, nPos, 1) <> \"/\"\n"
+ "\t\t\tnPos = nPos - 1\n"
+ "\t\tLoop\n"
+ "\t\tsMusic = Left (sFile, nPos) & sMusic\n"
+ "\tEnd If\n"
+ "\t\n"
+ "\tShell \"mplayer\", 6, \"\"\"\" & sMusic & \"\"\"\", False\n"
+ "End Sub\n";
    
    /** The BASIC macros to convert color numbers to mosaic art. */
    private static final String basicMacros =
"' Copyright (c) 2012 imacat\n"
+ "' \n"
+ "' Licensed under the Apache License, Version 2.0 (the \"License\");\n"
+ "' you may not use this file except in compliance with the License.\n"
+ "' You may obtain a copy of the License at\n"
+ "' \n"
+ "'     http://www.apache.org/licenses/LICENSE-2.0\n"
+ "' \n"
+ "' Unless required by applicable law or agreed to in writing, software\n"
+ "' distributed under the License is distributed on an \"AS IS\" BASIS,\n"
+ "' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.\n"
+ "' See the License for the specific language governing permissions and\n"
+ "' limitations under the License.\n"
+ "\n"
+ "' _1CalcMosaic: The BASIC macros of Calc Mosaic\n"
+ "'   Created with " + APP_NAME + " version " + VERSION + "\n"
+ "'   By imacat <imacat@mail.imacat.idv.tw>, 2012/10/5\n"
+ "\n"
+ "Option Explicit\n"
+ "\n"
+ "' subCreateMosaic: Creates the spreadsheets of mosaic art\n"
+ "Sub subCreateMosaic (mData () As Object)\n"
+ "\tDim mNames (UBound (mData (0))) As String\n"
+ "\tDim mColors (UBound (mData (1))) As Object\n"
+ "\tDim oSheet As Object, nI As Integer\n"
+ "\t\n"
+ "\t' Adds the spreadsheets and allocates memory first, to prevent\n"
+ "\t' memory corruption when the the number of sheets goes large.\n"
+ "\tmNames = mData (0)\n"
+ "\tmColors = mData (1)\n"
+ "\tsubInitializeSheets (mNames)\n"
+ "\tFor nI = 0 To UBound (mNames)\n"
+ "\t\tsubCreateMosaicSheet (mNames (nI), mColors (nI))\n"
+ "\tNext nI\n"
+ "\t\n"
+ "\tsubAddMacroTriggers\n"
+ "End Sub\n"
+ "\n"
+ "' subCreateMosaicSheet: Creates a spreadsheet of mosaic art with\n"
+ "'   the color values\n"
+ "Sub subCreateMosaicSheet (sSheet As String, mColors () As Object)\n"
+ "\tDim oSheet As Object, oRange As Object, oCell As Object\n"
+ "\tDim nRow As Integer, nColumn As Integer, nWidth As Integer\n"
+ "\t\n"
+ "\toSheet = ThisComponent.getSheets.getByName (sSheet)\n"
+ "\toRange = oSheet.getCellRangeByPosition ( _\n"
+ "\t\t0, 0, UBound (mColors (0)), UBound (mColors))\n"
+ "\tFor nRow = 0 To UBound (mColors)\n"
+ "\t\tFor nColumn = 0 To UBound (mColors (0))\n"
+ "\t\t\toCell = oRange.getCellByPosition (nColumn, nRow)\n"
+ "\t\t\toCell.setPropertyValue (\"CellBackColor\", _\n"
+ "\t\t\t\tCLng (mColors (nRow)(nColumn)))\n"
+ "\t\tNext nColumn\n"
+ "\tNext nRow\n"
+ "\t\n"
+ "\t' Shortly displays the sheet.  Removing these can save the\n"
+ "\t' display refreshing time and run slightly faster, with\n"
+ "\t' the cost of not seeing the progress and looking like hang.\n"
+ "\tnWidth = oRange.getRows.getPropertyValue (\"Height\")\n"
+ "\toRange.getColumns.setPropertyValue (\"Width\", nWidth)\n"
+ "\tThisComponent.getCurrentController.setActiveSheet (oSheet)\n"
+ "\tWait 1\n"
+ "End Sub\n"
+ "\n"
+ "' subInitializeSheets: Initialize the spreadsheets\n"
+ "Sub subInitializeSheets (mNames () As Object)\n"
+ "\tDim oSheets As Object, nI As Integer, mExisting () As String\n"
+ "\t\n"
+ "\toSheets = ThisComponent.getSheets\n"
+ "\tIf oSheets.hasByName (mNames (0)) Then\n"
+ "\t\tIf oSheets.getCount = 1 Then\n"
+ "\t\t\toSheets.insertNewByName (mNames (0) & 1, 0)\n"
+ "\t\tEnd If\n"
+ "\t\toSheets.removeByName (mNames (0))\n"
+ "\tEnd If\n"
+ "\tmExisting = oSheets.getElementNames\n"
+ "\toSheets.insertNewByName (mNames (0), 0)\n"
+ "\tFor nI = 0 To UBound (mExisting)\n"
+ "\t\toSheets.removeByName (mExisting (nI))\n"
+ "\tNext nI\n"
+ "\tFor nI = 1 To UBound (mNames)\n"
+ "\t\toSheets.insertNewByName (mNames (nI), nI)\n"
+ "\tNext nI\n"
+ "End Sub\n"
+ "\n"
+ "' fnGetEmptyStringArray: Returns an empty array of String.\n"
+ "Function fnGetEmptyStringArray (nUBound As Integer)\n"
+ "\tDim mArray (nUBound) As String\n"
+ "\t\n"
+ "\tfnGetEmptyStringArray = mArray\n"
+ "End Function\n"
+ "\n"
+ "' subAddMacroTriggers: Adds the macro trigggers to show the sheets\n"
+ "Sub subAddMacroTriggers\n"
+ "\tDim oSheets As Object, oSheet As Object\n"
+ "\tDim oCell As Object, nColor As Long\n"
+ "\t\n"
+ "\toSheets = ThisComponent.getSheets\n"
+ "\tsubAddStartButton (oSheets.getByIndex (0))\n"
+ "\t' Adds a start button at the end.  Last time (2012/8/18 on\n"
+ "\t' COSCUP 2012 lightning talk) I tested playing and ended at\n"
+ "\t' the last sheet.  When talk started I did not notice it and\n"
+ "\t' spent quite some time finding the button.  Audience got\n"
+ "\t' anxious and even lost attention when the show begins.\n" 
+ "\tsubAddStartButton (oSheets.getByIndex (oSheets.getCount - 1))\n"
+ "\t\n"
+ "\toSheet = ThisComponent.getSheets.getByIndex (0)\n"
+ "\tThisComponent.getCurrentController.setActiveSheet (oSheet)\n"
+ "\toCell = oSheet.getCellByPosition (1, 0)\n"
+ "\tnColor = oCell.getPropertyValue (\"CellBackColor\")\n"
+ "\toCell.setPropertyValue (\"CharColor\", nColor)\n"
+ "\toCell = oSheet.getCellByPosition (2, 0)\n"
+ "\tnColor = oCell.getPropertyValue (\"CellBackColor\")\n"
+ "\toCell.setPropertyValue (\"CharColor\", nColor)\n"
+ "End Sub\n"
+ "\n"
+ "' subAddStartButton: Adds a start button to show the sheets\n"
+ "Sub subAddStartButton (oSheet As Object)\n"
+ "\tDim oPage As Object, oShape As Object, oButton As Object\n"
+ "\tDim nSize As Integer, nColor As Long\n"
+ "\tDim aPos As New com.sun.star.awt.Point\n"
+ "\tDim aSize As New com.sun.star.awt.Size\n"
+ "\tDim oForm As Object, nIndex As Integer\n"
+ "\tDim aEvent As New com.sun.star.script.ScriptEventDescriptor\n"
+ "\t\n"
+ "\toPage = oSheet.getDrawPage\n"
+ "\tnSize = oSheet.getRows.getPropertyValue (\"Height\")\n"
+ "\t\n"
+ "\toShape = ThisComponent.createInstance ( _\n"
+ "\t\t\"com.sun.star.drawing.ControlShape\")\n"
+ "\taPos.X = 0\n"
+ "\taPos.Y = 0\n"
+ "\toShape.setPosition (aPos)\n"
+ "\taSize.Width = nSize\n"
+ "\taSize.Height = nSize\n"
+ "\toShape.setSize (aSize)\n"
+ "\t\n"
+ "\toButton = ThisComponent.createInstance ( _\n"
+ "\t\t\"com.sun.star.form.component.CommandButton\")\n"
+ "\tnColor = oSheet.getCellByPosition (0, 0).getPropertyValue ( _\n"
+ "\t\t\"CellBackColor\")\n"
+ "\toButton.setPropertyValue (\"BackgroundColor\", nColor)\n"
+ "\t\n"
+ "\toShape.setControl (oButton)\n"
+ "\toPage.add (oShape)\n"
+ "\t\n"
+ "\taEvent.ListenerType = \"XActionListener\"\n"
+ "\taEvent.EventMethod = \"actionPerformed\"\n"
+ "\taEvent.ScriptType = \"Script\"\n"
+ "\taEvent.ScriptCode = \"vnd.sun.star.script:\" _\n"
+ "\t\t& \"Standard." + playModule + ".subPlaySheets\" _\n"
+ "\t\t& \"?language=Basic&location=document\"\n"
+ "\toForm = oButton.getParent\n"
+ "\tnIndex = oForm.getCount - 1\n"
+ "\toForm.registerScriptEvent (nIndex, aEvent)\n"
+ "End Sub\n";
    
    /** The localization resources. */
    private ResourceBundle l10n = null;
    
    /** The window icon. */
    protected static Image icon = null;
    
    /** The image files. */
    private File files[] = new File[0];;
    
    /**
     * The number of rows to use, default to 0 (estimate it
     * automatically).
     */
    private int numRows = 0;
    
    /** Whether we should run in slow motion (for presentation). */
    private boolean isSlow = false;
    
    /** Should we uses GUI. */
    private boolean isGui = false;
    
    /** The spreadsheet document to work with. */
    private SpreadsheetDocument document = null;
    
    /** The available dimension of mosaic cells. */
    private Dimension availCells = null;
    
    /**
     * Creates a new instance of CalcMosaic.
     * 
     * @param args the command line arguments
     * @throws IllegalArgumentException if the arguments are
     *         illegal
     */
    private CalcMosaic(String args[])
            throws IllegalArgumentException {
        // Obtains the localization resources
        this.l10n = ResourceBundle.getBundle(
            this.getClass().getPackage().getName() + ".res.L10n");
        this.parseArguments(args);
        if (this.isGui) {
            icon = Toolkit.getDefaultToolkit().getImage(ICON);
            OptionDialog dialog = new OptionDialog(
                this.files, this.numRows, this.isSlow);
            Options options = dialog.getOptions();
            // Cancelled
            if (options == null) {
                System.exit(0);
            }
            this.files = options.files;
            this.numRows = options.numRows;
            this.isSlow = options.isSlow;
        }
    }
    
    /**
     * Creates a new instance of CalcMosaic, with the parameters.
     * 
     * @param files   The image files
     * @param numRows The number of rows to use, default to 0 (estimate it
     *                automatically)
     * @param isSlow  Should we create in slow motion (for presentation)
     */
    public CalcMosaic(File files[], int numRows, boolean isSlow) {
        // Obtains the localization resources
        this.l10n = ResourceBundle.getBundle(
            this.getClass().getPackage().getName() + ".res.L10n");
        this.files = files;
        this.numRows = numRows;
        this.isSlow = isSlow;
    }
    
    /**
     * Runs CalcMosaic from the command line.
     * 
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        CalcMosaic mosaic = null;
        
        try {
            mosaic = new CalcMosaic(args);
        } catch (IllegalArgumentException e) {
            System.err.println(APP_NAME + ": " + e.getMessage());
            System.exit(1);
        }
        try {
            mosaic.create();
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            mosaic.showErrorMessage(e.getMessage());
            System.exit(2);
        } catch (java.io.IOException e) {
            mosaic.showErrorMessage(e.getMessage());
            System.exit(2);
        } catch (NotAnImageException e) {
            mosaic.showErrorMessage(e.getMessage());
            System.exit(2);
        }
        mosaic.showMessage(mosaic._("finished"));
        System.exit(0);
    }
    
    /**
     * Calculates the mosaic color matrix from an image file.
     * 
     * @param image the image
     * @param cellSize the size of the mosaic cell
     * @return the values of the mosaic color matrix, with Alpha
     *         channel removed
     */
    public static int[][] addMacroTriggers(
            BufferedImage image, int cellSize) {
        int colors[][] = new int
            [(image.getHeight() - 1) / cellSize + 1]
            [(image.getWidth() - 1) / cellSize + 1];
        
        // Gets the color of each cell
        for (int y0 = 0; y0 < image.getHeight(); y0 += cellSize) {
            int y1 = y0 + cellSize - 1;
            if (y1 >= image.getHeight()) {
                y1 = image.getHeight() - 1;
            }
            for (int x0 = 0; x0 < image.getWidth(); x0 += cellSize) {
                int x1 = x0 + cellSize - 1;
                if (x1 >= image.getWidth()) {
                    x1 = image.getWidth() - 1;
                }
                Color color = getMosaicColor(image, x0, y0, x1, y1);
                colors[y0 / cellSize][x0 / cellSize]
                    = color.getRGB() & 0x00FFFFFF;
            }
        }
        
        return colors;
    }
    
    /**
     * Obtain the mosaic color of a block in the image.
     * 
     * @param image the image
     * @param x0 the x-coordinate of the upper-left corner of the
     *        block in the image
     * @param y0 the y-coordinate of the upper-left corner of the
     *        block in the image
     * @param x1 the x-coordinate of the lower-right corner of the
     *        block in the image
     * @param y1 the y-coordinate of the lower-right corner of the
     *        block in the image
     * @return the mosaic color of the block in the image
     */
    public static Color getMosaicColor(BufferedImage image,
            int x0, int y0, int x1, int y1) {
        /*
         * Using linear distance weakening:  When calculating the
         * average color of this block, the weights of the colors at
         * all the pixels are different:  Colors of pixels closer to
         * the center of the block have a larger weight that those
         * colors of pixels far away from center.  The weight
         * decreases from the center in a linear speed.  The resulting
         * weights may look like a circled cone.  This creates a
         * mosaic effect that is more sharp, especially with images of
         * only a few colors.  Our eye focus stay mostly at the center
         * of the block.  The colors of the center area outweights the
         * skirts area.
         */
        
        double totalWeight = 0;
        double alpha = 0, red = 0, green = 0, blue = 0;
        double centerX = (double) (x0 + x1) / 2;
        double centerY = (double) (y0 + y1) / 2;
        double radius = 1 + Math.sqrt(
            (x0 - centerX) * (x0 - centerX)
            + (y0 - centerY) * (y0 - centerY));
        
        for (int x = x0; x <= x1; x++) {
            for (int y = y0; y <= y1; y++) {
                Color color = new Color(image.getRGB(x, y));
                double distance = Math.sqrt(
                    (x - centerX) * (x - centerX)
                    + (y - centerY) * (y - centerY));
                double weight = (radius - distance) / radius;
                
                alpha += color.getAlpha() * weight;
                red += color.getRed() * weight;
                green += color.getGreen() * weight;
                blue += color.getBlue() * weight;
                totalWeight += weight;
            }
        }
        alpha /= totalWeight;
        red /= totalWeight;
        green /= totalWeight;
        blue /= totalWeight;
        
        return new Color((int) red, (int) green, (int) blue,
            (int) alpha);
    }
    
    /**
     * Parses the command line arguments.
     * 
     * @param args the command line arguments
     * @throws IllegalArgumentException if the arguments are
     *         illegal
     */
    private void parseArguments(String args[])
            throws IllegalArgumentException {
        List<String> filesList = new ArrayList<String>();
        
        for (int i = 0; i < args.length; i++) {
            if (args[i].equals("-r") || args[i].equals("-rows")) {
                i++;
                if (i >= args.length) {
                    throw new NumberFormatException(
                        _("err_rows_miss"));
                }
                try {
                    this.numRows = Integer.parseInt(args[i]);
                } catch (NumberFormatException e) {
                    throw new NumberFormatException(
                        args[i - 1] + " \"" + args[i] + "\": "
                            + _("err_rows_not_pint"));
                }
                if (this.numRows <= 0) {
                    throw new NumberFormatException(
                        args[i - 1] + " \"" + args[i] + "\": "
                            + _("err_rows_not_pint"));
                }
            } else if (args[i].startsWith("-r=")) {
                try {
                    this.numRows = Integer.parseInt(args[i].substring(3));
                } catch (NumberFormatException e) {
                    throw new NumberFormatException(
                        "\"" + args[i] + "\": "
                            + _("err_rows_not_pint"));
                }
                if (this.numRows <= 0) {
                    throw new NumberFormatException(
                        "\"" + args[i] + "\": "
                            + _("err_rows_not_pint"));
                }
            } else if (args[i].startsWith("-rows=")) {
                try {
                    this.numRows = Integer.parseInt(args[i].substring(6));
                } catch (NumberFormatException e) {
                    throw new NumberFormatException(
                        "\"" + args[i] + "\": "
                            + _("err_rows_not_pint"));
                }
                if (this.numRows <= 0) {
                    throw new NumberFormatException(
                        "\"" + args[i] + "\": "
                            + _("err_rows_not_pint"));
                }
            } else if (args[i].equals("-s") || args[i].equals("-slow")) {
                this.isSlow = true;
            } else if (args[i].equals("-noslow")) {
                this.isSlow = false;
            } else if (args[i].equals("-h") || args[i].equals("-help")) {
                System.out.println(String.format(_("help"), JAR_NAME));
                System.exit(0);
            } else if (args[i].equals("-v") || args[i].equals("-version")) {
                System.out.println(String.format(_("version"),
                    APP_NAME, VERSION, _("author")));
                System.exit(0);
            } else if (args[i].startsWith("-")) {
                throw new IllegalArgumentException(
                    args[i] + ": " + _("err_arg_unknown"));
            } else {
                filesList.add(args[i]);
            }
        }
        
        // Checks the files.
        // This is processed last, after all the options, so that
        // -help and -version takes effect before this.
        this.files = new File[filesList.size()];
        for (int i = 0; i < this.files.length; i++) {
            File file = new File(filesList.get(i));
            
            if (!file.exists()) {
                throw new IllegalArgumentException(
                    args[i] + ": " + _("err_file_not_exist"));
            }
            if (!file.isFile()) {
                throw new IllegalArgumentException(
                    args[i] + ": " + _("err_file_not_a_file"));
            }
            this.files[i] = file;
        }
        // Checks the limits of the number of images/spreadsheets.
        if (this.files.length > 256) {
            throw new IllegalArgumentException(
                _("err_file_too_many"));
        }
        
        // Use GUI if no files are supplied.
        if (this.files.length == 0) {
            this.isGui = true;
        }
    }
    
    /**
     * Creates mosaic art of images with Calc.
     * 
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws java.io.IOException if an error occurs during reading
     *         the image
     * @throws NotAnImageException if the file read is not an image
     */
    private void create()
            throws com.sun.star.comp.helper.BootstrapException,
                java.io.IOException,
                NotAnImageException {
        if (this.isSlow) {
            this.createSlowly();
        } else {
            this.createNormal();
        }
    }
    
    /**
     * Creates mosaic art of images with Calc with the normal speed.
     * 
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws java.io.IOException if an error occurs during reading
     *         the image
     * @throws NotAnImageException if the file read is not an image
     */
    private void createNormal()
            throws com.sun.star.comp.helper.BootstrapException,
                java.io.IOException,
                NotAnImageException {
        OpenOffice office = null;
        Object param[] = new Object[1];
        Object data[][] = new Object[2][this.files.length];
        
        try {
            office = new OpenOffice();
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            throw e;
        }
        this.document = office.startNewSpreadsheetDocument();
        this.document.addBasicModule(playModule, playMacros);
        this.document.addBasicModule(basicModule, basicMacros);
        for (int i = 0; i < this.files.length; i++) {
            data[0][i] = this.files[i].getName();
            try {
                data[1][i] = this.addMacroTriggers(this.files[i]);
            } catch (java.io.IOException e) {
                throw e;
            } catch (NotAnImageException e) {
                throw e;
            }
        }
        param[0] = data;
        this.document.runBasicMacro(
            basicModule, "subCreateMosaic", param);
    }
    
    /**
     * Creates mosaic art of images with Calc in slow motion.
     * 
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws java.io.IOException if an error occurs during reading
     *         the image
     * @throws NotAnImageException if the file read is not an image
     */
    private void createSlowly()
            throws com.sun.star.comp.helper.BootstrapException,
                java.io.IOException,
                NotAnImageException {
        OpenOffice office = null;
        
        try {
            office = new OpenOffice();
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            throw e;
        }
        this.document = office.startNewSpreadsheetDocument();
        this.document.addBasicModule(playModule, playMacros);
        for (int i = 0; i < this.files.length; i++) {
            String name = this.files[i].getName();
            int colors[][] = null;
            
            try {
                colors = this.addMacroTriggers(this.files[i]);
            } catch (java.io.IOException e) {
                throw e;
            } catch (NotAnImageException e) {
                throw e;
            }
            this.document.createMosaicSheet(name, i, colors);
        }
        this.document.addMacroTriggers();
    }
    
    /**
     * Calculates the mosaic color matrix from an image file.
     * 
     * @param file the source image file
     * @return the values of the mosaic color matrix, with Alpha
     *         channel removed
     * @throws java.io.IOException if an error occurs during reading
     *         the image
     * @throws NotAnImageException if the file read is not an image
     */
    private int[][] addMacroTriggers(File file)
            throws java.io.IOException,
                NotAnImageException {
        BufferedImage image = null;
        int cellSize = -1;
        
        try {
            image = ImageIO.read(file);
        } catch (java.io.IOException e) {
            throw e;
        }
        if (image == null) {
            throw new NotAnImageException(
                file.getName() + ": " + _("err_file_not_an_image"));
        }
        
        cellSize = this.getCellSize(image);
        return addMacroTriggers(image, cellSize);
    }
    
    /**
     * Returns the mosaic cell size.
     * 
     * @param image the source image
     * @return the calculated mosaic cell size
     */
    private int getCellSize(BufferedImage image) {
        if (this.numRows != 0) {
            return (image.getHeight() - 1) / this.numRows + 1;
        }
        
        // Calculate the size that fits the whole image.
        int width = 0, height = 0;
        
        this.estimateAvailableCells();
        width = (image.getWidth() - 1) / this.availCells.width + 1;
        height = (image.getHeight() - 1) / this.availCells.height + 1;
        return width > height? width: height;
    }
    
    /**
     * Estimates the number of rows to use automatically
     * 
     * @param image the source image
     * @return the number of rows to use
     */
    private int estimateNumRows(BufferedImage image) {
        // Calculate the size that fits the whole image.
        int width = -1, height = -1, cellSize = -1;
        
        this.estimateAvailableCells();
        width = (image.getWidth() - 1) / this.availCells.width + 1;
        height = (image.getHeight() - 1) / this.availCells.height + 1;
        cellSize = width > height? width: height;
        return Math.round((float) image.getHeight() / cellSize);
    }
    
    /**
     * Estimates the available dimension of the mosaic cells.
     * 
     */
    private void estimateAvailableCells() {
        if (this.availCells == null) {
            Dimension window = this.document.getWindowSize();
            int rowHeight = this.document.getRowHeight();
            
            // Only use the row height, because the height is always
            // shorter than the width in newly-created sheets.
            this.availCells = new Dimension(
                window.width * 25 / rowHeight,
                window.height * 25 / rowHeight);
        }
    }
    
    /**
     * Shows a error message.
     * 
     * @param message the error message to be shown
     */
    private void showErrorMessage(String message) {
        if (this.isGui) {
            JFrame frame = new JFrame(APP_NAME);
            Dimension screenSize
                = Toolkit.getDefaultToolkit().getScreenSize();
            Dimension frameSize = frame.getSize();
            
            frame.setIconImage(icon);
            frame.setLocation(
                    (screenSize.width - frameSize.width) / 2,
                    (screenSize.height - frameSize.height) / 2
                );
            frame.setVisible(true);
            JOptionPane.showMessageDialog(frame, message,
                String.format(_("dialog_err_title"),
                    CalcMosaic.APP_NAME),
                JOptionPane.ERROR_MESSAGE);
            frame.setVisible(false);
        } else {
            System.err.println(APP_NAME + ": " + message);
        }
    }
    
    /**
     * Shows a message.
     * 
     * @param message the message to be shown
     */
    private void showMessage(String message) {
        if (this.isGui) {
            JFrame frame = new JFrame(APP_NAME);
            Dimension screenSize
                = Toolkit.getDefaultToolkit().getScreenSize();
            Dimension frameSize = frame.getSize();
            
            frame.setIconImage(icon);
            frame.setLocation(
                    (screenSize.width - frameSize.width) / 2,
                    (screenSize.height - frameSize.height) / 2
                );
            frame.setVisible(true);
            JOptionPane.showMessageDialog(frame, message, APP_NAME,
                JOptionPane.INFORMATION_MESSAGE);
            frame.setVisible(false);
        } else {
            System.out.println(APP_NAME + ": " + message);
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
                l10n.getString(key).getBytes("ISO-8859-1"), "UTF-8");
        } catch (java.io.UnsupportedEncodingException e) {
            return l10n.getString(key);
        } catch (java.util.MissingResourceException e) {
            return key;
        }
    }
}
