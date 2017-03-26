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
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.artofsolving.jodconverter.util.PlatformUtils;

import com.sun.star.beans.PropertyValue;
import com.sun.star.uno.UnoRuntime;

public class OfficeUtils {

	public static final String SERVICE_DESKTOP = "com.sun.star.frame.Desktop";

	public static final String[] WINDOWS_INSTALLATION_LOCATIONS = { //
			"OpenOffice 4", //
			"LibreOffice 4", //
			"OpenOffice.org 3", //
			"LibreOffice 3", //
	};

	public static final String[] MACOS_INSTALLATION_LOCATIONS = { //
			"/Applications/OpenOffice.app/Contents", // V4+
			"/Applications/LibreOffice.app/Contents", // V$+
			"/Applications/OpenOffice.org.app/Contents", // V3
	};

	public static final String[] UNIX_INSTALLATION_LOCATIONS = { //
			"/opt/libreoffice", //
			"/usr/lib/openoffice", //
			"/usr/lib/libreoffice", //
			"/opt/openoffice.org3", // V3
	};

	private static final Logger logger = Logger.getLogger(OfficeUtils.class.getName());

	private OfficeUtils() {
		throw new AssertionError("utility class must not be instantiated");
	}

	public static <T> T cast(Class<T> type, Object object) {
		return (T) UnoRuntime.queryInterface(type, object);
	}

	public static PropertyValue property(String name, Object value) {
		PropertyValue propertyValue = new PropertyValue();
		propertyValue.Name = name;
		propertyValue.Value = value;
		return propertyValue;
	}

	@SuppressWarnings("unchecked")
	public static PropertyValue[] toUnoProperties(Map<String, ?> properties) {
		PropertyValue[] propertyValues = new PropertyValue[properties.size()];
		int i = 0;
		for (Map.Entry<String, ?> entry : properties.entrySet()) {
			Object value = entry.getValue();
			if (value instanceof Map) {
				Map<String, Object> subProperties = (Map<String, Object>) value;
				value = toUnoProperties(subProperties);
			}
			propertyValues[i++] = property((String) entry.getKey(), value);
		}
		return propertyValues;
	}

	public static String toUrl(File file) {
		String path = file.toURI().getRawPath();
		String url = path.startsWith("//") ? "file:" + path : "file://" + path;
		return url.endsWith("/") ? url.substring(0, url.length() - 1) : url;
	}

	public static File getDefaultOfficeHome() {
		if (System.getProperty("office.home") != null) {
			return new File(System.getProperty("office.home"));
		}

		return findOfficeHome(findPossibleOfficeHomeLocations());
	}

	public static String[] findPossibleOfficeHomeLocations() {
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

	private static File findOfficeHome(String... knownPaths) {
		for (String path : knownPaths) {
			File home = new File(path);
			if (getOfficeExecutable(home).isFile()) {
				return home;
			}
		}
		return null;
	}

	public static File getOfficeExecutable(File officeHome) {
		if (PlatformUtils.isMac()) {
			File sofficeExecutable = new File(officeHome, "MacOS/soffice.bin");

			// LibreOffice doesn't have soffice.bin, just soffice
			if (!sofficeExecutable.exists())
				sofficeExecutable = new File(officeHome, "MacOS/soffice");

			logger.info("Found: " + sofficeExecutable);
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
}
