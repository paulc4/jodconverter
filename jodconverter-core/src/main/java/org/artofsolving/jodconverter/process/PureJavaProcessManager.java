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

import java.io.IOException;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.util.PlatformUtils;

public class PureJavaProcessManager implements ProcessManager {

	public long findPid(ProcessQuery query) {
		return PID_UNKNOWN;
	}

	public void kill(Process process, long pid) {

		if (process == null) {
			String processId = String.valueOf(pid);
			if (PlatformUtils.isWindows())
				execute(pid, "taskkill", "/pid", processId, "/f");
			else
				execute(pid, "/bin/kill", "-KILL", processId);
		} else {
			process.destroy();
		}
	}

	private List<String> execute(long pid, String... args) {
		try {
			String[] command = args;
			Process process = new ProcessBuilder(command).start();
			@SuppressWarnings("unchecked")
			List<String> lines = IOUtils.readLines(process.getInputStream());
			return lines;
		} catch (IOException e) {
			throw new OfficeException("Error trying to kill process " + pid);
		}
	}
}
