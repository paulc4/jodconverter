//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter.office;

import static org.artofsolving.jodconverter.office.OfficeUtils.toUrl;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import java.io.File;
import java.util.List;

import org.artofsolving.OfficeSoftware;
import org.junit.Before;
import org.junit.Test;

public class OfficeUtilsTest {

	private boolean hasLibreOffice = false;
	private boolean hasOpenOffice = false;

	public static void main(String[] args) {
		OfficeUtilsTest tester = new OfficeUtilsTest();
		tester.setup();
		tester.testFindExecutable();
	}

	@Before
	public void setup() {
		List<File> executables = OfficeUtils.findAllExecutables();

		for (File f : executables) {
			if (f.toString().contains("Libre"))
				hasLibreOffice = true;
			else if (f.toString().contains("Open"))
				hasOpenOffice = true;
		}
	}

	@Test
	public void testToUrl() {
		// TODO create separate tests for Windows
		assertEquals(toUrl(new File("/tmp/document.odt")), "file:///tmp/document.odt");
		assertEquals(toUrl(new File("/tmp/document with spaces.odt")), "file:///tmp/document%20with%20spaces.odt");
	}

	@Test
	public void testFindExecutable() {
		testFindExecutable(OfficeSoftware.LIBRE_OFFICE, hasLibreOffice);
		testFindExecutable(OfficeSoftware.OPEN_OFFICE, hasOpenOffice);
	}

	private static void testFindExecutable(OfficeSoftware whichOffice, boolean exists) {

		File executable = OfficeUtils.getOfficeExecutable(whichOffice);
		System.out.println("FOUND: " + executable);

		if (exists)
			assertNotNull(executable);
		else
			assertNull(executable);
	}

}
