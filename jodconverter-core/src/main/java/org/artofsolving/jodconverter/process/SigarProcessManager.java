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
package org.artofsolving.jodconverter.process;

import java.io.File;
import java.io.IOException;
import java.util.logging.Logger;

import org.artofsolving.jodconverter.util.ZipUtils;
import org.hyperic.sigar.Sigar;
import org.hyperic.sigar.SigarException;
import org.hyperic.sigar.ptql.ProcessFinder;

/**
 * {@link ProcessManager} implementation that uses the SIGAR library.
 * <p>
 * Requires the sigar.jar in the classpath and the appropriate system-specific
 * native library (e.g. <tt>libsigar-x86-linux.so</tt> on Linux x86) available
 * in the <em>java.library.path</em>.
 * <p>
 * See the <a href="http://support.hyperic.com/display/SIGAR">SIGAR site</a> for
 * documentation and downloads.
 */
public class SigarProcessManager implements ProcessManager {

	public static final String SIGAR_NATIVE_LIB_ZIP = "sigar-native-lib.zip";
	public static final String SIGAR_FALLBACK_LIB_DIR = ".jod-sigar-lib";

	public static final String PATH_SEPARATOR = "path.separator";
	public static final String JAVA_LIBRARY_PATH = "java.library.path";

	// If no library directory is explicitly specified, default to
	// ~/.jod-sigar-lib
	private static final String HOME_DIR = System.getProperty("user.home");
	private static final String SIGAR_NATIVE_LIBRARY_PATH = HOME_DIR + '/' + SIGAR_FALLBACK_LIB_DIR;

	private static boolean setLibraryPath = false;

	private static final Logger logger = Logger.getLogger(SigarProcessManager.class.getName());

	/**
	 * Downloaded SIGAR from https://sourceforge.net/projects/sigar and zipped
	 * up the native libraries/DLLs into:
	 * <tt>src/main/resources/sigar-native-lib.zip</tt>.
	 * <p>
	 * They need to be on the Java library path to be picked-up. They are added
	 * at the end so if a different version has been added already, it will be
	 * found first.
	 * <p>
	 * <b>Note:</b> SIGAR (System Information Gatherer and Reporter) is a
	 * cross-platform, cross-language library and command-line tool for
	 * accessing operating system and hardware level information in Java, Perl
	 * and .NET.
	 * <p>
	 * Originally developed by Hyperic, the company was taken over by VMware.
	 */
	public static void setupLibraryPath() {
		if (setLibraryPath)
			return;

		String actualLibDir = SIGAR_NATIVE_LIBRARY_PATH + "/sigar-native-lib";

		// Have we unpacked the zip before?
		File sigarJar = new File(actualLibDir + "/sigar.jar");

		if (sigarJar.exists() && sigarJar.isFile())
			; // Nothing to do
		else {
			// Unpack the zip included in this project
			ZipUtils.unzipToUserHome(SIGAR_NATIVE_LIB_ZIP, SIGAR_FALLBACK_LIB_DIR, ZipUtils.DONT_OVERWRITE);
			new File(SIGAR_FALLBACK_LIB_DIR, "__MACOSX").delete();
		}

		setupLibraryPath(actualLibDir);
	}

	/**
	 * Explicitly set the location of the SIGAR native jars. They can be found
	 * at https://sourceforge.net/projects/sigar. Once set, this method quietly
	 * does nothing if called again.
	 * 
	 * <b>Note:</b> SIGAR (System Information Gatherer and Reporter) is a
	 * cross-platform, cross-language library and command-line tool for
	 * accessing operating system and hardware level information in Java, Perl
	 * and .NET.
	 * <p>
	 * Originally developed by Hyperic, the company was taken over by VMware.
	 * 
	 * @param sigarNativeLibDir
	 *            Wherever you have copied the native libraries to.
	 */
	public static void setupLibraryPath(String sigarNativeLibDir) {
		if (setLibraryPath)
			return;

		String libPath = System.getProperty(JAVA_LIBRARY_PATH);
		String sep = System.getProperty(PATH_SEPARATOR);
		System.setProperty(JAVA_LIBRARY_PATH, libPath + sep + sigarNativeLibDir);

		logger.info(System.getProperty(JAVA_LIBRARY_PATH));
		setLibraryPath = true;
	}

	public SigarProcessManager() {
		setupLibraryPath();
	}

	public long findPid(ProcessQuery query) throws IOException {
		Sigar sigar = new Sigar();
		logger.info("Looking for process " + query.getCommand() + " " + query.getArgument());
		try {
			long[] pids = ProcessFinder.find(sigar, "State.Name.eq=" + query.getCommand());
			for (int i = 0; i < pids.length; i++) {
				String[] arguments = sigar.getProcArgs(pids[i]);
				if (arguments != null && argumentMatches(arguments, query.getArgument())) {
					logger.info("Process " + query.getCommand() + " found: id=" + pids[i]);
					return pids[i];
				}
			}

			logger.info("Process " + query.getCommand() + " bot found");
			return PID_NOT_FOUND;
		} catch (SigarException sigarException) {
			throw new IOException("findPid failed", sigarException);
		} finally {
			sigar.close();
		}
	}

	public void kill(Process process, long pid) throws IOException {
		Sigar sigar = new Sigar();

		// Developer error?
		if (pid <= 0) {
			return;
		}

		try {
			logger.info("kill " + pid);
			sigar.kill(pid, Sigar.getSigNum("KILL"));
		} catch (SigarException sigarException) {
			throw new IOException("kill " + pid + ": failed", sigarException);
		} finally {
			sigar.close();
		}
	}

	private boolean argumentMatches(String[] arguments, String expected) {
		for (String argument : arguments) {
			if (argument.contains(expected)) {
				return true;
			}
		}
		return false;
	}

}
