package org.artofsolving;

import java.io.File;
import java.io.IOException;
import java.util.Date;
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
import org.artofsolving.jodconverter.office.OfficeUtils;
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
	 * What to do if the output file exists already. The default is to just
	 * overwrite it anyway (FORCE).
	 */
	public enum OverwritePolicy {
		/** Don't overwrite and throw an error */
		FAIL,
		/** Don't overwrite and return quietly */
		SKIP,
		/** Overwrite regardless */
		FORCE,
		/** Overwrite if the input is newer than the output */
		NEWER
	}

	/**
	 * Default server port when Open/Libre Office runs as a server.
	 */
	public static final int DEFAULT_PORT = 8100;

	// Protected data members to allow sub-classing
	protected final Map<String, Object> outputProperties = new HashMap<String, Object>();
	protected final DefaultOfficeManagerConfiguration configuration;
	protected final Logger logger;

	protected OfficeManager officeManager = null;
	protected OfficeDocumentConverter documentConverter;
	protected OfficeSoftware preferred;
	protected OverwritePolicy overwritePolicy = OverwritePolicy.FORCE;

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

	/**
	 * Create an instance using all the defaults. Will make smart guesses to try
	 * and find Open or Libre Office assuming you have installed them in their
	 * default location. The specify port number is used.
	 * <p>
	 * Server will be available on <tt>localhost:[port]</tt>. The communication
	 * uses URP (UNO remote protocol). UNO stands for Universal Network Objects,
	 * and is the component model of OpenOffice.org.
	 * 
	 * @param officePort
	 *            The port Open/Libre Office will listen on.
	 */
	public OpenOfficeConverterUsingJOD3(int officePort) {
		this(officePort, "", "");
	}

	public OpenOfficeConverterUsingJOD3(int officePort, String officeHome, String officeProfile) {

		configuration = new DefaultOfficeManagerConfiguration();
		configuration.setPortNumber(officePort);

		System.setProperty(DefaultOfficeManagerConfiguration.KEEP_PROFILE_DIRECTORY, "false");

		if (hasText(officeHome))
			configuration.setOfficeHome(new File(officeHome));
		if (hasText(officeProfile))
			configuration.setTemplateProfileDir(new File(officeProfile));

		logger = Logger.getLogger(getClass().getName());
	}

	/**
	 * Which software is preferred? Open or Libre Office. Open Office works best
	 * with Open Office, only Libre Office handles the Microsoft so-called Open
	 * XML Formats (docx, pptx, xlsx).
	 * 
	 * @param whichOffice
	 *            Which one to use - or {@link OfficeSoftware#NONE} if you don't
	 *            care.
	 */
	public void setPreference(OfficeSoftware whichOffice) {
		this.preferred = whichOffice;
		OfficeUtils.setPreferredSoftware(whichOffice);
	}

	/**
	 * Which software is preferred? Open or Libre Office. Open Office works best
	 * with Open Office, only Libre Office handles the Microsoft so-called Open
	 * XML Formats (docx, pptx, xlsx). If the requested software cannot be
	 * found, an exception is thrown.
	 * 
	 * @param whichOffice
	 *            Which one to use - or {@link OfficeSoftware#NONE} if you don't
	 *            care.
	 */
	public void setMandatoryPreference(OfficeSoftware whichOffice) {
		this.preferred = whichOffice;
		OfficeUtils.setMandatoryPreferredSoftware(whichOffice);
	}

	/**
	 * What to do if the output file already exists. Default is to overwrite
	 * (FORCE). Options are to just ignore the input, leaving output file
	 * unchanged (SKIP), throw an error (FAIL) or to overwrite if the input file
	 * is newer than the existing output file. (NEWER).
	 * 
	 * @param overwritePolicy
	 */
	public void setOverwritePolicy(OverwritePolicy overwritePolicy) {
		this.overwritePolicy = overwritePolicy;
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
		File workDir = new File(System.getProperty("user.home"), ".common-build");
		workDir.mkdir();
		configuration.setWorkDir(workDir);

		if (officeManager == null) {
			officeManager = configuration.buildOfficeManager();
			documentConverter = new OfficeDocumentConverter(officeManager);
		}

		officeManager.start();
	}

	public void stop() {
		if (officeManager == null)
			throw new IllegalStateException("Office manager not started, nothing to stop");
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
	 * called {@link #start()} before invoking this method. Don't forget to
	 * invoke {@link #stop()} when you have finished. If you only need to
	 * convert a single file, {@link #convertFileOneOff(File, File)} is more
	 * convenient.
	 * <p>
	 * This method works in one of two ways:
	 * <ul>
	 * <li>If <tt>destination</tt> is an existing directory, the method will put
	 * the newly created file into that directory. The new file will have the
	 * same name as the input file with a different suffix to reflect the output
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

		// What to do if output file exists already
		if (outputFile.exists()) {
			switch (overwritePolicy) {
			case FAIL:
				throw new OfficeException("conversion failed - output file exists already");
			case FORCE:
				logger.info(outputFile + " exists - overwrite");
				// Nothing to do
				break;
			case NEWER:
				logger.info("Compare: IN " + inputFile.lastModified() + " OUT " + outputFile.lastModified());
				logger.info("Compare: IN " + new Date(inputFile.lastModified()) + " OUT "
						+ new Date(outputFile.lastModified()));
				if (inputFile.lastModified() < outputFile.lastModified()) {
					logger.info(outputFile + " exists and is newer - skip");
					return outputFile;
				}

				logger.info(outputFile + " exists but is older - overwrite");
				break;
			case SKIP:
				logger.info(outputFile + " exists - skip");
				return outputFile;
			}
		} else {
			logger.info("CREATING NEW FILE: " + outputFile);
		}

		boolean startedByMe = false;
		long startTime = System.currentTimeMillis();
		Exception exception = null;
		int retries = 0;
		int maxRetries = 2;

		while (retries++ < maxRetries) {
			if (retries > 1)
				logger.info("About to convert: attempt #" + retries);

			try {
				if (!officeManager.isRunning()) {
					logger.info("Starting soffice");
					officeManager.start();
					startedByMe = true;
				} else if (retries > 1)
					logger.info("soffice is running");

				DocumentFormat outputFormat = converter.getFormatRegistry().getFormatByExtension(outputExtension);
				logger.fine(" >>> OutputFormat = " + outputFormat);

				if (!outputProperties.isEmpty()) {
					outputFormat.setStoreProperties(DocumentFamily.PRESENTATION, outputProperties);
				}

				logger.fine("Output Properties: " + outputFormat.getStoreProperties(DocumentFamily.PRESENTATION));
				// System.out.println(outputFormat.getExtension());
				// System.out.println(outputFormat.getMediaType());
				// Run the converter

				converter.convert(inputFile, outputFile, outputFormat);
				long conversionTime = System.currentTimeMillis() - startTime;
				logger.info(String.format("successful conversion: %s [%dKb] to %s in %dms", inputExtension,
						(inputFile.length() / 1000), outputExtension, conversionTime));

				return outputFile;
			} catch (Exception e) {
				long timeoutAfter = System.currentTimeMillis() - startTime;
				logger.severe(String.format("failed conversion %s [%dKb] to %s  after %dms: %s", inputFile.getName(),
						(inputFile.length() / 1000), outputExtension, timeoutAfter, e));
				exception = e;
			} finally {
				if (startedByMe) {
					logger.info("Stopping soffice");
					officeManager.stop();
				}
			}
		}

		throw new OfficeException("conversion failed", exception);
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
