package org.artofsolving;

import java.io.File;
import java.util.List;

import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeUtils;
import org.junit.Test;

public class ConversionToPDFTests {

	public static final int DEFAULT_PORT = 8100;

	public static final String TEST_CONTENT = //
			"src/test/resources/documents/test.";

	public static final File TEST_INTRO_PPTX = //
			new File("src/test/resources/01-Introduction.pptx");

	public static final File TEST_INTRO_PPT = //
			new File("src/test/resources/00-Introduction.ppt");

	public static final File TEST_INTRO_ODP = //
			new File("src/test/resources/sw-introduction.odp");

	public static final File[] TEST_INPUT_FILES = { //
			TEST_INTRO_ODP, //
			TEST_INTRO_PPT, //
			TEST_INTRO_PPTX, //
	};

	@Test
	public void findOfficeHomeLocations() {
		List<File> executables = OfficeUtils.findAllExecutables();
		System.out.println(executables);
	}

	@Test
	public void odpToPdfUsingDefaults() {
		OpenOfficeConverterUsingJOD3 converter = new OpenOfficeConverterUsingJOD3();

		File destinationDir = new File("target/test-output");
		destinationDir.mkdirs();

		converter.start();
		converter.convertFile(TEST_INTRO_ODP, destinationDir, "PDF");
		converter.stop();
	}

	@Test
	public void filesToPdf() {
		// String officeHome = "/Applications/OpenOffice.app/Contents";
		String officeHome = "/Applications/LibreOffice.app/Contents";

		String userHome = System.getenv("HOME");
		// String officeProfile = userHome + //
		// "/Library/Application Support/OpenOffice/4";
		String officeProfile = userHome + "/Library/Application Support/LibreOffice/4";

		OpenOfficeConverterUsingJOD3 converter = new OpenOfficeConverterUsingJOD3(DEFAULT_PORT, officeHome,
				officeProfile);

		File destinationDir = new File("target/test-output");
		destinationDir.mkdirs();

		converter.start();

		for (File inputFile : TEST_INPUT_FILES)
			converter.convertFile(inputFile, destinationDir, "PDF");

		converter.stop();
	}

	@Test
	public void odpToPdfNotes() {
		DefaultOfficeManagerConfiguration.usingProxySoSetRealPort(8100);

		OpenOfficeConverterUsingJOD3 converter = new OpenOfficeConverterUsingJOD3(8111, "", "");

		File input = new File(TEST_CONTENT + "odp");
		File destination = new File("target/test-output/notes.pdf");
		destination.getParentFile().mkdirs();

		converter.setOutputProperty("ExportNotesPages", "true"); // Boolean.TRUE);
		converter.convertFileOneOff(input, destination);
	}
}
