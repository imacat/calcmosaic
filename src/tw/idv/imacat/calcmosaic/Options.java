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
 * Options
 * 
 * Created on 2012-10-29
 * 
 * Copyright (c) 2012 imacat
 */

package tw.idv.imacat.calcmosaic;

import java.io.File;

/**
 * The options of Calc Mosaic.
 * 
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 0.0.1
 */
public class Options {
    
    /** The image files. */
    public File files[] = new File[0];;
    
    /** The number of rows to use. */
    public int numRows = 0;
    
    /** Whether we should run in slow motion (for presentation). */
    public boolean isSlow = false;
    
    /**
     * Creates a new set of options.
     * 
     * @param files   the image files
     * @param numRows the number of rows to use
     * @param isSlow  whether we should run in slow motion (for
     *                presentation)
     */
    public Options(File files[], int numRows, boolean isSlow) {
        this.files = files;
        this.numRows = numRows;
        this.isSlow = isSlow;
    }
}
