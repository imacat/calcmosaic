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
 * NotAnImageException
 * 
 * Created on 2012-10-27, rewritten from CalcMosaic
 * 
 * Copyright (c) 2012 imacat
 */

package tw.idv.imacat.calcmosaic;

/**
 * An exception thrown when the file read is not an image.
 * 
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 0.0.1
 */
public class NotAnImageException extends Exception {
    
    /**
     * Constructs a <code>NotAnImageException</code> exception with
     * <code>null</code> as its error detail message.
     * 
     */
    public NotAnImageException() {
        super();
    }
    
    /**
     * Constructs a <code>NotAnImageException</code> with the
     * specified detail message.
     * 
     * @param message The detail message (which is saved for later
     *                retrieval by the Throwable.getMessage() method)
     */
    public NotAnImageException(String message) {
        super(message);
    }
    
    /**
     * Constructs an <code>NotAnImageException</code> with the
     * specified detail message and cause. 
     * <p>
     * Note that the detail message associated with <code>cause</code>
     * is <em>not</em> automatically incorporated into this
     * exception's detail message.
     * 
     * @param message The detail message (which is saved for later
     *                retrieval by the Throwable.getMessage() method)
     * @param cause   The cause (which is saved for later retrieval
     *                by the Throwable.getCause() method). (A null
     *                value is permitted, and indicates that the
     *                cause is nonexistent or unknown.)
     */
    public NotAnImageException(String message, Throwable cause) {
        super(message, cause);
    }
    
    /**
     * Constructs an <code>NotAnImageException</code> with the
     * specified cause and a detail message of <code>(cause==null
     * &#63; null : cause.toString())</code> (which typically contains
     * the class and detail message of <code>cause</code>).  This
     * constructor is useful for IO exceptions that are little more
     * than wrappers for other throwables.
     * 
     * @param cause   The cause (which is saved for later retrieval
     *                by the Throwable.getCause() method). (A null
     *                value is permitted, and indicates that the
     *                cause is nonexistent or unknown.)
     */
    public NotAnImageException(Throwable cause) {
        super(cause);
    }
}
