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
 * OpenOffice
 * 
 * Created on 2012-08-16, rewritten from CalcMosaic
 * 
 * Copyright (c) 2012 imacat
 */

package tw.idv.imacat.calcmosaic;

import java.util.ResourceBundle;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * An OpenOffice instance.
 * 
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 2.0.0
 */
public class OpenOffice {
    
    /** The localization resources. */
    private ResourceBundle l10n = null;
    
    /** The desktop service. */
    private XDesktop desktop = null;
    
    /** The bootstrap context. */
    private XComponentContext bootstrapContext = null;
    
    /** The registry service manager. */
    private XMultiComponentFactory serviceManager = null;
    
    /**
     * Creates a new instance of Office.
     * 
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     */
    public OpenOffice()
            throws com.sun.star.comp.helper.BootstrapException {
        // Obtains the localization resources
        this.l10n = ResourceBundle.getBundle(
            this.getClass().getPackage().getName() + ".res.L10n");
        try {
            this.connect();
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            throw e;
        }
    }
    
    /**
     * Connects to the OpenOffice program.
     * 
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     */
    private void connect()
            throws com.sun.star.comp.helper.BootstrapException {
        XMultiComponentFactory localServiceManager = null;
        Object unoUrlResolver = null;
        XUnoUrlResolver xUnoUrlResolver = null;
        Object service = null;
        
        if (this.isConnected()) {
            return;
        }
        
        // Obtains the local context
        try {
            this.bootstrapContext = Bootstrap.bootstrap();
        } catch (java.lang.Exception e) {
            throw new com.sun.star.comp.helper.BootstrapException(e);
        }
        if (this.bootstrapContext == null) {
            throw new com.sun.star.comp.helper.BootstrapException(
                this._("err_ooo_no_lcc"));
        }
        
        // Obtains the local service manager
        this.serviceManager = this.bootstrapContext.getServiceManager();
        
        // Obtain the desktop service
        try {
            service = this.serviceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", this.bootstrapContext);
        } catch (com.sun.star.uno.Exception e) {
            throw new java.lang.UnsupportedOperationException(e);
        }
        this.desktop = (XDesktop) UnoRuntime.queryInterface(
            XDesktop.class, service);
        
        // Obtains the service manager
        this.serviceManager = this.bootstrapContext.getServiceManager();
        return;
    }
    
    /**
     * Returns whether the connection is on and alive.
     * 
     * @return true if the connection is alive, false otherwise
     */
    private boolean isConnected() {
        if (this.serviceManager == null) {
            return false;
        }
        try {
            UnoRuntime.queryInterface(
                XPropertySet.class, this.serviceManager);
        } catch (com.sun.star.lang.DisposedException e) {
            this.serviceManager = null;
            return false;
        }
        return true;
    }
    
    /**
     * Starts a new spreadsheet document.
     * 
     * @return the new spreadsheet document
     */
    public SpreadsheetDocument startNewSpreadsheetDocument() {
        XComponentLoader xComponentLoader = (XComponentLoader)
            UnoRuntime.queryInterface(
            XComponentLoader.class, this.desktop);
        PropertyValue props[] = new PropertyValue[0];
        final String url = "private:factory/scalc";
        XComponent xComponent = null;
        
        // Load a new document
        try {
            xComponent = xComponentLoader.loadComponentFromURL(url,
                "_default", 0, props);
        } catch (com.sun.star.io.IOException e) {
            throw new java.lang.RuntimeException(
                String.format(_("err_open"), url), e);
        } catch (com.sun.star.lang.IllegalArgumentException e) {
            throw new java.lang.RuntimeException(
                String.format(_("err_open"), url), e);
        }
        
        return new SpreadsheetDocument(xComponent);
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
