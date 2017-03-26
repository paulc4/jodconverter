package org.artofsolving;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.commons.io.FilenameUtils;
import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.artofsolving.jodconverter.process.SigarProcessManager;

/**
 * File format converter using Open or Libre Office running in remote server
 * mode. For convenience this class is all you need (hides all other JOD
 * Converter classes).
 * <p>
 * Because this runs a remote server you need to invoke methods in the right
 * order:
 * 
 * <pre>
 * OpenOfficeConverter converter = new OpenOfficeConverter();
 * 
 * // Specify location of SIGAR libraries 
 * converter.setSigarNativeLibraryPath("/path/to/sigar/native/libs");
 * 
 * // Start the server
 * converter.start();
 * 
 * // Creates mypresso.pdf in /tmp
 * coverter.convert(new File("mypresso.odp", new File("/tmp"), "PDF");     
 * 
 * // Creates presso2.pdf
 * coverter.convert(new File("presso2.ppt", new File("presso2.pdf")) ;  
 * 
 * // Do as many conversions as you need.  To only run one conversion
 * // consider {@link #convertFileOneOff(File, File)} instead (you still
 * // need to specify the SIGAR library directory).
 * 
 * // Stop the server
 * converter.stop();
 * </pre>
 * 
 * @author Paul Chapman
 * @see https://sourceforge.net/projects/sigar
 */
// In the accept string: urp = UNO remote protocol
public class OpenOfficeConverterUsingJOD3 {

	/**
	 * Default server port when Open/Libre Office runs as a server.
	 */
	public static final int DEFAULT_PORT = 8100;

	// Protected data members to allow sub-classing
	protected final OfficeManager officeManager;
	protected final OfficeDocumentConverter documentConverter;
	protected final Map<String, Object> outputProperties = new HashMap<String, Object>();
	protected final Logger logger;

	/**
	 * Create an instance using all the defaults. Will make smart guesses to try
	 * and find Open or Libre Office assuming you have installed them in their
	 * default location. The default port (8100) is used.
	 * <p>
	 * Server will be available on <tt>localhost:8100</tt>. The communication
	 * uses URP (UNO remote protocol). UNO stands for Universal Network Objects,
	 * and is the component model of OpenOffice.org.
	 */
	public OpenOfficeConverterUsingJOD3() {
		this(DEFAULT_PORT, "", "");
	}

	public OpenOfficeConverterUsingJOD3(int officePort, String officeHome, String officeProfile) {

		DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
		configuration.setPortNumber(officePort);

		if (hasText(officeHome))
			configuration.setOfficeHome(new File(officeHome));
		if (hasText(officeProfile))
			configuration.setTemplateProfileDir(new File(officeProfile));

		officeManager = configuration.buildOfficeManager();
		documentConverter = new OfficeDocumentConverter(officeManager);
		logger = Logger.getLogger(getClass().getName());
	}

	/**
	 * Uses SIGAR internally to automatically start and stop open office. SIGAR
	 * relies on Operating System functionality and native libraries/DLLS to do
	 * this. Set the location if you have your preferred version.
	 * 
	 * @param sigarNativeLibDir
	 *            Location of SIGAR native libraries and DLLs.
	 * 
	 * @see https://sourceforge.net/projects/sigar
	 */
	public void setSigarNativeLibraryPath(String sigarNativeLibDir) {
		SigarProcessManager.setupLibraryPath(sigarNativeLibDir);
	}

	/**
	 * Specify a property that can be used to configure how the output is
	 * generated. See the URL below for a list.
	 * 
	 * @param name
	 * @param value
	 * 
	 * @see https://wiki.openoffice.org/wiki/API/Tutorials/PDF_export
	 */
	public void setOutputProperty(String name, Object value) {
		outputProperties.put(name, value);
	}

	public void clearOutputProperties() {
		outputProperties.clear();
	}

	/**
	 * Start Open or Libre Office as a background process (it runs in headless
	 * server mode).
	 */
	public void start() {
		officeManager.start();
	}

	public void stop() {
		officeManager.stop();
	}

	/**
	 * Convert a single file as a one-off task. This starts and stops the
	 * OpenOffice or LibreOffice executable which takes time. to convert
	 * multiple files, use {@link #convertFile(File, File, String)}.
	 * 
	 * @param inputFile
	 *            An OpenOffice or Microsoft Office file.
	 * @param destinationFile
	 *            The file to create. The file suffix will be used to determine
	 *            the output format fir the conversion.
	 * @return The newly created file - <tt>destinationFile</tt>
	 */
	public File convertFileOneOff(File inputFile, File destinationFile) {
		try {
			start();
			String outputType = FilenameUtils.getExtension(destinationFile.getName());
			return convertFile(inputFile, destinationFile, outputType);
		} finally {
			stop();
		}
	}

	/**
	 * Convert the input file into the specified output format. You must have
	 * called {@link #start()} before invoking this method. Don't foget to
	 * invoke {@link #stop()} when you have finished. If you only need to
	 * convert a single file, {@link #convertFileOneOff(File, File)} is more
	 * convenient.
	 * <p>
	 * This method works in one of two ways:
	 * <ul>
	 * <li>If <tt>destination</tt> is an existing directory, the method will put
	 * the newly created file into that directory. The new file will have the
	 * ame name as the input file with a different suffix to reflect the output
	 * type.
	 * <li>Otherwise <tt>destination</tt> is assumed to be the file to generate.
	 * It's parent directory must exist.
	 * 
	 * @param inputFile
	 * @param destination
	 * @param outputType
	 * @return
	 */
	public File convertFile(File inputFile, File destination, String outputType) {

		checkAvailable(inputFile);
		OfficeDocumentConverter converter = this.documentConverter;

		String inputExtension = FilenameUtils.getExtension(inputFile.getName());
		String outputExtension = outputType.toLowerCase();

		File destinationDir;
		File outputFile;

		if (destination.isDirectory()) {
			destinationDir = destination;
			String baseName = FilenameUtils.getBaseName(inputFile.getName());
			outputFile = new File(destinationDir, baseName + "." + outputExtension);
		} else {
			outputFile = destination;
			destinationDir = destination.getParentFile();

			if (!destinationDir.exists())
				throw new OfficeException(
						destination.toString() + " cannot be created because its parent directory does not exist");
		}

		boolean startedByMe = false;

		try {
			if (!officeManager.isRunning()) {
				officeManager.start();
				startedByMe = true;
			}

			DocumentFormat outputFormat = converter.getFormatRegistry().getFormatByExtension(outputExtension);

			if (!outputProperties.isEmpty()) {
				outputFormat.setStoreProperties(DocumentFamily.PRESENTATION, outputProperties);
			}

			logger.info("Output Properties: " + outputFormat.getStoreProperties(DocumentFamily.PRESENTATION));

			// Run the converter
			long startTime = System.currentTimeMillis();
			converter.convert(inputFile, outputFile, outputFormat);
			long conversionTime = System.currentTimeMillis() - startTime;
			logger.info(String.format("successful conversion: %s [%db] to %s in %dms", inputExtension,
					inputFile.length(), outputExtension, conversionTime));

			return outputFile;
		} catch (Exception exception) {
			logger.severe(String.format("failed conversion: %s [%db] to %s; %s; input file: %s", inputExtension,
					inputFile.length(), outputExtension, exception, inputFile.getName()));
			throw new OfficeException("conversion failed", exception);
		} finally {
			if (startedByMe)
				officeManager.stop();
		}
	}

	// A COUPLE OF UTILITY METHODS

	/**
	 * Does the specified string contain some text?
	 * 
	 * @param s
	 *            Any string
	 * @return False if <tt>s is null or the empty string, true otherwise.
	 */
	protected boolean hasText(String s) {
		return s != null && s.length() > 0;
	}

	/**
	 * Check if a file or directory is available for use (does it exist?).
	 * 
	 * @param file
	 *            A file-system resource.
	 * @throws IOException
	 * @throws IllegalArgumentException
	 *             Test failed.: file not found.
	 */
	public static void checkAvailable(File file) {
		if (!file.exists()) {
			throw new IllegalArgumentException("File or directory '" + file + "' does not exist");
		}
	}

}
