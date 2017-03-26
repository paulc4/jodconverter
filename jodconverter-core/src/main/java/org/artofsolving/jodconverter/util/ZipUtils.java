package org.artofsolving.jodconverter.util;

import java.io.BufferedOutputStream;
import java.io.Closeable;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.logging.Logger;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.artofsolving.jodconverter.office.OfficeException;

import java.lang.IllegalArgumentException;

/**
 * Convenient utility class for extracting files from a Zip.
 * <p>
 * Code courtesy of Sunli Kumar Sahoo
 *
 * @author Paul Chapman
 * @see http://stackoverflow.com/questions/1705844/what-is-the-best-way-to-extract-a-zip-file-using-java
 */
public class ZipUtils {

	/**
	 * Default directory when using - {@value}.
	 */
	static public final String DEFAULT_TEMP_DIRECTORY = ".jod-temp";

	static public final boolean OVERWRITE = true;
	static public final boolean DONT_OVERWRITE = false;

	static private final String homeDir = System.getProperty("user.home");
	static private Logger logger = Logger.getLogger(ZipUtils.class.getName());

	/**
	 * Unpack to a new temporary directory. If a temporary directory cannot be
	 * created, it will fail.
	 * 
	 * @param zipResource
	 *            A Zip file on the classpath
	 * @return The new location.
	 * @throws OfficeException
	 *             Unable to create temporary directory.
	 */
	public static File unzipToTempdir(String zipResource) {
		File destinationDir;

		try {
			destinationDir = Files.createTempDirectory("jod-temp").toFile();
		} catch (IOException e) {
			logger.severe("Unable to create temporary directory: " + e);
			throw new OfficeException("Unable to create temporary directory", e);
		}

		unzip(zipResource, destinationDir);
		return destinationDir;
	}

	/**
	 * Unpack to a temporary directory. If a temporary directory cannot be
	 * created an exception will be raised. If it already exists and overwrite
	 * is false nothing is unpacked and this method quietly quits. (the fallback
	 * directory is .
	 * 
	 * @param zipResource
	 *            A Zip file on the classpath
	 * @param dirName
	 *            Name of directory to create under home directory. If this is
	 *            null or the empty string then {@link #DEFAULT_TEMP_DIRECTORY}
	 *            is used.
	 * @param overwrite
	 *            If the directory exists already, overwrite or not?
	 * @return The new location.
	 */
	public static File unzipToUserHome(String zipResource, String dirName, boolean overwrite) {
		File destinationDir;

		if (dirName == null || dirName.length() == 0)
			dirName = DEFAULT_TEMP_DIRECTORY;

		String destinationPath = new StringBuffer().append(homeDir).append(File.separator).append(dirName).toString();
		destinationDir = new File(destinationPath);

		if (!destinationDir.exists()) {
			boolean ok = destinationDir.mkdirs();

			if (!ok)
				logger.warning("Unable to create directory " + destinationDir);
		} else if (destinationDir.isFile()) {
			String error = "Unable to create directory (a file with that name already exists): " + destinationDir;
			logger.severe(error);
			throw new OfficeException(error);
		} else if (!overwrite) {
			return destinationDir; // nothing to do
		}

		unzip(zipResource, destinationDir);
		logger.info("Unpacked " + zipResource + " into" + dirName);
		return destinationDir;
	}

	/**
	 * Unpack into the specified directory.
	 * 
	 * @param zipResource
	 *            A Zip file on the classpath
	 * @param destinationDir
	 *            A directory to unpack into - it must exist
	 * @throws IllegalArgumentException
	 *             If <tt>destinationDir</tt> does not exist.
	 */
	public static void unzip(String zipResource, File destinationDir) {
		int BUFFER = 10000;

		if (!destinationDir.exists())
			throw new IllegalArgumentException("ERROR: " + destinationDir + "does not exist");

		try {
			BufferedOutputStream out = null;
			ZipInputStream in = new ZipInputStream(ZipUtils.class.getClassLoader().getResourceAsStream(zipResource));
			ZipEntry entry;
			boolean isDirectory = false;

			while ((entry = in.getNextEntry()) != null) {
				int count;
				byte data[] = new byte[BUFFER];

				// Write the files to the disk
				String entryName = entry.getName();
				File newFile = new File(destinationDir, entryName);

				if (entryName.endsWith("/")) {
					isDirectory = true;
					newFile.mkdir();
					// System.out.println("This is directory
					// "+newFile.exists()+" IS DIr "+newFile.isDirectory()+"
					// path "+newFile.getPath());
				} else {
					newFile.createNewFile();
				}

				if (!isDirectory) {
					out = new BufferedOutputStream(new FileOutputStream(newFile), BUFFER);
					while ((count = in.read(data, 0, BUFFER)) != -1) {
						out.write(data, 0, count);
					}
					close(out);
				}

				isDirectory = false;
			}

			close(in);
		} catch (Exception e) {
			e.printStackTrace();
			System.exit(0);
		}
	}

	/**
	 * Convenient close resource method (catches the useless IOException that
	 * might be thrown).
	 * 
	 * @param stream
	 *            Any {@link Closeable} input or output.
	 */
	private static void close(Closeable stream) {
		try {
			stream.close();
		} catch (IOException e) {
			// What to do here? I mean really?
		}
	}

}
