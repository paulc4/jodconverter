package org.artofsolving.jodconverter.util;

import java.io.File;

import org.junit.Assert;
import org.junit.Test;

public class ZipUtilsTests {

	public static final String SIGAR_NATIVE_LIB_ZIP = "sigar-native-lib.zip";

	// 24 Libraries/DLLs/Jars plus .sigar_shellrc
	int NUM_LIBS_EXPECTED = 25;

	@Test
	public void extractSigarLibrariesToTemp() {
		String zip = SIGAR_NATIVE_LIB_ZIP;

		File destinationDir = ZipUtils.unzipToTempdir(zip);
		validateResults(destinationDir);
	}

	@Test
	public void extractSigarLibrariesToLocal() {
		String zip = SIGAR_NATIVE_LIB_ZIP;

		File destinationDir = ZipUtils.unzipToUserHome(zip, "jod-temp", ZipUtils.OVERWRITE);
		validateResults(destinationDir);
	}

	protected void validateResults(File destinationDir) {
		System.out.println("Created " + destinationDir);

		File libDir = new File(destinationDir, "sigar-native-lib");

		String[] libs = libDir.list();
		Assert.assertEquals("Expected " + NUM_LIBS_EXPECTED + " files, but found " //
				+ libs.length, NUM_LIBS_EXPECTED, libs.length);

		boolean foundSigarJar = false;

		for (String lib : libs) {
			foundSigarJar |= lib.equals("sigar.jar");
		}

		Assert.assertTrue("sigar.jar not found in " + destinationDir, foundSigarJar);
	}
}
