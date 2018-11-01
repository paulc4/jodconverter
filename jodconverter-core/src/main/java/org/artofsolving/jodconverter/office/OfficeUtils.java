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

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.logging.Logger;

import org.artofsolving.OfficeSoftware;
import org.artofsolving.jodconverter.util.PlatformUtils;

import com.sun.star.beans.PropertyState;
import com.sun.star.beans.PropertyValue;
import com.sun.star.uno.UnoRuntime;

public class OfficeUtils {

	private static final String LIBRE = "Libre";

	public static final String SERVICE_DESKTOP = "com.sun.star.frame.Desktop";

	public static final int LATEST_VERSION = 4;

	public static final int MINIMAL_ACCEPTABLE_VERSION = 3;

	public static final int ANY_VERSION = 0;

	public static final String[] WINDOWS_INSTALLATION_LOCATIONS = { //
			"OpenOffice 4", //
			"LibreOffice 4", //
			"OpenOffice.org 3", //
			"LibreOffice 3", //
	};

	public static final String[] MACOS_INSTALLATION_LOCATIONS = { //
			"/Applications/LibreOffice 2.app/Contents", // V5.3+
			"/Applications/OpenOffice.app/Contents", // V4+
			"/Applications/LibreOffice.app/Contents", // V4+
			"/Applications/OpenOffice.org.app/Contents", // V3
	};

	public static final String[] UNIX_INSTALLATION_LOCATIONS = { //
			"/opt/libreoffice", //
			"/usr/lib/openoffice", //
			"/usr/lib/libreoffice", //
			"/opt/openoffice.org3", // V3
	};

	private static final Logger logger = Logger.getLogger(OfficeUtils.class.getName());

	private static OfficeSoftware preferred = OfficeSoftware.NONE;

	private static boolean isPreferenceMandatory = false;

	private static final List<String> versionNumbers = new ArrayList<>();

	// Populate with possible version numbers, plus 0 (version unknown)
	static {
		for (int version = LATEST_VERSION; version >= MINIMAL_ACCEPTABLE_VERSION; version--)
			versionNumbers.add(String.valueOf(version));

		versionNumbers.add(String.valueOf(ANY_VERSION));
	}

	private OfficeUtils() {
		throw new AssertionError("utility class must not be instantiated");
	}

	// --------------------------------------------------------------------//
	// - - - - - - - - - - - - - - UNO PROTOCOL - - - - - - - - - - - - - //
	// --------------------------------------------------------------------//

	public static <T> T cast(Class<T> type, Object object) {
		return (T) UnoRuntime.queryInterface(type, object);
	}

	public static PropertyValue property(String name, Object value) {
		PropertyValue propertyValue = new PropertyValue();
		propertyValue.Name = name;
		propertyValue.Value = value;
		propertyValue.Handle = -1;

		String type = (propertyValue.State == PropertyState.DIRECT_VALUE) ? "DIRECT"
				: (propertyValue.State == PropertyState.DEFAULT_VALUE ? "DEFAULT" : "AMBIGUOUS");
		logger.fine("PropertyValue: " + propertyValue.Name + " " + //
				propertyValue.Handle + " " + propertyValue.Value + //
				" (" + propertyValue.Value.getClass() + ") " + type);

		return propertyValue;
	}

	@SuppressWarnings("unchecked")
	public static PropertyValue[] toUnoProperties(Map<String, ?> properties) {
		PropertyValue[] propertyValues = new PropertyValue[properties.size()];
		int i = 0;

		if (properties.containsKey("ReadOnly"))
			i++;

		for (Map.Entry<String, ?> entry : properties.entrySet()) {
			Object value = entry.getValue();
			if (value instanceof Map) {
				Map<String, Object> subProperties = (Map<String, Object>) value;
				value = toUnoProperties(subProperties);
			}
			PropertyValue pv = property((String) entry.getKey(), value);

			if ("ReadOnly".equals(pv.Name))
				propertyValues[0] = pv;
			else
				propertyValues[i++] = pv;
		}
		return propertyValues;
	}

	public static String toUrl(File file) {
		String path = file.toURI().getRawPath();
		String url = path.startsWith("//") ? "file:" + path : "file://" + path;
		return url.endsWith("/") ? url.substring(0, url.length() - 1) : url;
	}

	// --------------------------------------------------------------------//
	// - - - - - - - - - - - OPEN OFFICE EXECUTABLE - - - - - - - - - - - //
	// --------------------------------------------------------------------//

	public static void setMandatoryPreferredSoftware(OfficeSoftware whichOffice) {
		preferred = whichOffice;
		isPreferenceMandatory = true;
	}

	public static void setPreferredSoftware(OfficeSoftware whichOffice) {
		preferred = whichOffice;
	}

	public static File getDefaultOfficeHome() {
		if (System.getProperty("office.home") != null) {
			return new File(System.getProperty("office.home"));
		}

		File home = findOfficeHome(findPossibleOfficeHomeLocations());
		logger.info(" >>> Office home: " + home);
		return home;
	}

	private static String[] findPossibleOfficeHomeLocations() {
		if (PlatformUtils.isWindows()) {
			// %ProgramFiles(x86)% on 64-bit machines; %ProgramFiles% on 32-bit
			// ones
			String programFiles = System.getenv("ProgramFiles(x86)");
			if (programFiles == null) {
				programFiles = System.getenv("ProgramFiles");
			}

			// Build a list of possible location within programFiles
			int i = 0;
			String[] locationPaths = new String[WINDOWS_INSTALLATION_LOCATIONS.length];

			for (String location : WINDOWS_INSTALLATION_LOCATIONS) {
				locationPaths[i++] = programFiles + File.separator + location;
			}

			return locationPaths;
		} else if (PlatformUtils.isMac()) {
			return MACOS_INSTALLATION_LOCATIONS;
		} else {
			// Linux or other *nix variants
			return UNIX_INSTALLATION_LOCATIONS;
		}
	}

	/**
	 * Looks for the location of the Office Executable. If
	 * 
	 * @param knownPaths
	 * @return
	 */
	private static File findOfficeHome(String... knownPaths) {
		Map<String, File> possibles = new HashMap<>();

		for (String path : knownPaths) {
			File home = new File(path);
			File executable = getOfficeExecutable(home);
			if (executable.isFile()) {
				String type = path.contains(LIBRE) ? OfficeSoftware.LIBRE_OFFICE.toString()
						: OfficeSoftware.OPEN_OFFICE.toString();
				String version = getVersion(path, executable);
				possibles.put(type + '-' + version, home);
			}
		}

		logger.info("Preferred = " + preferred);
		logger.info("Possibles = \n    " + possibles);

		// Nothing found
		if (possibles.isEmpty())
			return null;

		// Try to find the latest version of the preferred software
		if (preferred != OfficeSoftware.NONE) {
			// Count backwards through the versions
			for (String versionId : versionNumbers) {
				String key = preferred.toString() + '-' + versionId;
				File home = possibles.get(key);

				if (home != null)
					return home;
			}

			// If the preference is mandatory, throw an exception
			if (isPreferenceMandatory)
				throw new SoftwareNotFoundException("Unable to find " + preferred + //
						" - set path explicitly using \"os.home\" System property", preferred);
		}

		File first = null;

		// Try to find the latest version
		for (String versionId : versionNumbers) {
			for (Map.Entry<String, File> f : possibles.entrySet()) {
				if (f.getKey().contains(versionId))
					return f.getValue();
				else if (first == null)
					first = f.getValue();
			}
		}

		// Give up - accept first file found
		return first;
	}

	private static String getVersion(String path, File executable) {
		if (PlatformUtils.isWindows()) {
			// Windows paths contain version number
			for (int version = LATEST_VERSION; version >= MINIMAL_ACCEPTABLE_VERSION; version--) {
				String versionId = String.valueOf(version);
				if (path.contains(versionId))
					return versionId;
			}
		}

		// MacOS and Linux, look for versionrc
		File versionrc = new File(executable.getParentFile(), "versionrc");
		logger.fine("Trying: " + versionrc);
		String versionId = String.valueOf(ANY_VERSION);

		if (!versionrc.exists()) {
			versionrc = new File(path, "/Resources/versionrc");

			if (!versionrc.exists())
				return versionId;
		}

		Properties props = new Properties();

		try {
			props.load(new FileInputStream(versionrc));
		} catch (IOException e) {
			e.printStackTrace();
		}

		logger.fine("    " + props);
		String versionProperty = props.getProperty("ProductMajor");

		if (versionProperty == null)
			versionProperty = props.getProperty("ReferenceOOoMajorMinor");

		logger.fine("    Version property = " + versionProperty);

		if (versionProperty != null && versionProperty.length() > 0)
			versionId = versionProperty.substring(0, 1);

		return versionId;
	}

	public static File getOfficeExecutable(File officeHome) {
		if (PlatformUtils.isMac()) {
			File sofficeExecutable = new File(officeHome, "MacOS/soffice.bin");

			// LibreOffice doesn't have soffice.bin, just soffice
			if (!sofficeExecutable.exists())
				sofficeExecutable = new File(officeHome, "MacOS/soffice");

			logger.fine("Found: " + sofficeExecutable);
			return sofficeExecutable;
		} else {
			return new File(officeHome, "program/soffice.bin");
		}
	}

	public static List<File> findAllExecutables() {
		String[] knownPaths = findPossibleOfficeHomeLocations();
		List<File> results = new ArrayList<File>();

		for (String path : knownPaths) {
			File home = new File(path);
			if (getOfficeExecutable(home).isFile()) {
				results.add(home);
			}
		}

		return results;
	}

	public static File getOfficeExecutable(OfficeSoftware whichOffice) {
		setPreferredSoftware(whichOffice);
		return getOfficeExecutable(getDefaultOfficeHome());
	}
}
