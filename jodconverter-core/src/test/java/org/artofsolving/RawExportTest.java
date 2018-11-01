package org.artofsolving;

import static org.artofsolving.jodconverter.office.OfficeUtils.SERVICE_DESKTOP;

import java.util.Arrays;

import org.artofsolving.jodconverter.office.OfficeConnection;
import org.artofsolving.jodconverter.office.OfficeUtils;
import org.artofsolving.jodconverter.office.UnoUrl;
import org.junit.Ignore;
import org.junit.Test;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lib.uno.typeinfo.MethodTypeInfo;
import com.sun.star.lib.uno.typeinfo.TypeInfo;
import com.sun.star.text.XText;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 * This doesn't work.
 */
public class RawExportTest {

	@Test
	@Ignore
	public void test() throws Exception {

		OfficeConnection conn = new OfficeConnection(UnoUrl.socket(8100));
		conn.connect();

		XComponent xComponent = OfficeUtils.cast(XComponent.class, conn.getService(SERVICE_DESKTOP));
		System.out.println(xComponent);
		// create a new Writer document using a helper
		// xComponent = Helper.createNewDoc(m_xContext, "swriter");

		// access its XTextDocument interface
		XTextDocument xTextDocument = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xComponent);
		System.out.println(xTextDocument);

		for (TypeInfo t : Arrays.asList(XComponent.UNOTYPEINFO)) {
			if (t instanceof MethodTypeInfo)
				System.out.println("   " + ((MethodTypeInfo) t).getName());
			else
				System.out.println("   " + t.getClass());
		}

		// access the text body and set a string
		XText xText = xTextDocument.getText();
		xText.setString("Simple PDF export demo.");

		// XStorable to store the document
		XStorable xStorable = (XStorable) UnoRuntime.queryInterface(XStorable.class, xComponent);

		// XStorable.storeToURL() expects an URL telling where to store the
		// document
		// and an array of PropertyValue indicating how to store it

		// URL = use helper method to get the home directory
		String sURL = /* Helper.getHomeDir(m_xContext) + */ "target/test-output/Simple_PDF_EXPORT_demo.pdf";

		// Exporting to PDF consists of giving the proper
		// filter name in the property "FilterName"
		// With only this, the document will be exported
		// using the existing PDF export settings
		// (the one used the last time, or the default if the first time)
		PropertyValue[] aMediaDescriptor = new PropertyValue[1];
		aMediaDescriptor[0] = new PropertyValue();
		aMediaDescriptor[0].Name = "FilterName";
		aMediaDescriptor[0].Value = "writer_pdf_Export";

		xStorable.storeToURL(sURL, aMediaDescriptor);
	}
}
