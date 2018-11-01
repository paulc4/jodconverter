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

import static org.artofsolving.jodconverter.process.ProcessManager.PID_NOT_FOUND;
import static org.artofsolving.jodconverter.process.ProcessManager.PID_UNKNOWN;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.commons.io.FileUtils;
import org.artofsolving.jodconverter.process.ProcessManager;
import org.artofsolving.jodconverter.process.ProcessQuery;
import org.artofsolving.jodconverter.util.PlatformUtils;

class OfficeProcess {

	/**
	 * If using a proxy, we need to to use the actual port. To modify use
	 * {@link DefaultOfficeManagerConfiguration#usingProxySoSetRealPort(int)}.
	 */
	public static int realServerPort = -1;

	/**
	 * Controls start behavior.
	 */
	public enum StartAction {
		/**
		 * Only start if not already running - raise exception otherwise
		 */
		IF_NOT_RUNNING,

		/**
		 * Always start a new server - restart if one is already running.
		 */
		RESTART,

		/**
		 * Start if not server is running. Quietly do nothing otherwise.
		 */
		UNLESS_RUNNING
	}

	private final Logger logger;
	private final File officeHome;
	private final UnoUrl unoUrl;
	private final String[] runAsArgs;
	private final File templateProfileDir;
	private final File instanceProfileDir;
	private final ProcessManager processManager;

	private File executable;
	private Process process;
	private long pid = PID_UNKNOWN;
	private boolean firstTime = true;
	private boolean deleteProfDir = true;

	public OfficeProcess(File officeHome, UnoUrl unoUrl, String[] runAsArgs, File templateProfileDir, File workDir,
			ProcessManager processManager) {
		this.logger = Logger.getLogger(getClass().getName());
		this.officeHome = officeHome;

		// If a proxy is being use, we need to use the real port number. Socket
		// connections only.
		this.unoUrl = realServerPort == -1 && unoUrl.getAcceptString().startsWith("socket") //
				? unoUrl : UnoUrl.socket(realServerPort);
		this.runAsArgs = runAsArgs;
		this.templateProfileDir = templateProfileDir;
		this.instanceProfileDir = getInstanceProfileDir(workDir, unoUrl);
		this.processManager = processManager;
	}

	/**
	 * Start a new server - will fail if there is one running already.
	 * 
	 * @throws IOException
	 */
	public void start() throws IOException {
		start(StartAction.IF_NOT_RUNNING);
	}

	/**
	 * Start a new server - will fail if there is one running already
	 * <i>unless</i> restart is requested.
	 * 
	 * @param restart
	 * @throws IOException
	 */
	public void start(boolean restart) throws IOException {
		start(restart ? StartAction.RESTART : StartAction.IF_NOT_RUNNING);
	}

	/**
	 * Start an Open/Libre Office server as a background headless process.
	 * 
	 * @param startAction
	 *            Startup options.
	 * @throws IOException
	 */
	public void start(StartAction startAction) throws IOException {
		executable = OfficeUtils.getOfficeExecutable(officeHome);

		if (firstTime) {
			logger.info("Server socket: " + unoUrl);
			logger.info("Server executable: " + executable);
		}

		// Is there an exiting process? If so fail unless restart specified.
		ProcessQuery processQuery = new ProcessQuery(executable.getName(), unoUrl.getAcceptString());
		long existingPid = processManager.findPid(processQuery);
		boolean isRunning = existingPid > 0;

		switch (startAction) {
		case IF_NOT_RUNNING:
			if (isRunning)
				throw new IllegalStateException(
						String.format("a process with acceptString '%s' is already running; pid %d",
								unoUrl.getAcceptString(), existingPid));
		case RESTART:
			if (isRunning)
				processManager.kill(process, existingPid);
			break;
		case UNLESS_RUNNING:
			if (isRunning) {
				logger.info("Process already running, using it");
				return;
			}
		}

		// Clear up any previous directory
		if (startAction != StartAction.RESTART) {
			prepareInstanceProfileDir();
		}

		List<String> command = new ArrayList<String>();
		if (runAsArgs != null) {
			command.addAll(Arrays.asList(runAsArgs));
		}

		// LibreOffice uses Unix-style --options and the command is different
		// (the --env option stops it starting)
		boolean usingLibreOffice = executable.toString().contains("Libre");
		String option = usingLibreOffice ? "--" : "-";

		command.add(executable.getAbsolutePath());
		command.add(option + "accept=" + unoUrl.getAcceptString() + ";urp;");
		command.add(option + "headless");

		if (!usingLibreOffice) {
			command.add(option + "env:UserInstallation=" + OfficeUtils.toUrl(instanceProfileDir));
		}

		command.add(option + "nocrashreport");
		command.add(option + "nodefault");
		command.add(option + "nofirststartwizard");
		command.add(option + "nolockcheck");
		command.add(option + "nologo");
		command.add(option + "norestore");

		if (firstTime)
			logger.info("Running command=" + command);

		ProcessBuilder processBuilder = new ProcessBuilder(command);
		if (PlatformUtils.isWindows()) {
			addBasisAndUrePaths(processBuilder);
		}

		if (firstTime)
			logger.info(String.format("starting process with acceptString '%s' and profileDir '%s'", unoUrl,
					instanceProfileDir));

		process = processBuilder.start();
		pid = processManager.findPid(processQuery);

		if (pid == PID_NOT_FOUND) {
			throw new IllegalStateException(String.format(
					"process with acceptString '%s' started but its pid could not be found", unoUrl.getAcceptString()));
		}

		if (firstTime)
			logger.info("started process" + (pid != PID_UNKNOWN ? "; pid = " + pid : ""));

		firstTime = false;
	}

	private File getInstanceProfileDir(File workDir, UnoUrl unoUrl) {
		String dirName = ".jodconverter_" + unoUrl.getAcceptString().replace(',', '_').replace('=', '-');
		return new File(workDir, dirName);
	}

	private void prepareInstanceProfileDir() throws OfficeException {
		deleteProfDir = Boolean.parseBoolean(System.getProperty( //
				DefaultOfficeManagerConfiguration.KEEP_PROFILE_DIRECTORY, "true"));

		if (instanceProfileDir.exists()) {
			if (deleteProfDir) {
				logger.warning(String.format("profile dir '%s' already exists; deleting", instanceProfileDir));
				deleteProfileDir();
			}
		}
		if (templateProfileDir != null) {
			try {
				FileUtils.copyDirectory(templateProfileDir, instanceProfileDir);
			} catch (IOException ioException) {
				throw new OfficeException("failed to create profileDir", ioException);
			}
		}
	}

	public void deleteProfileDir() {
		if (instanceProfileDir != null) {
			try {
				if (deleteProfDir)
					FileUtils.deleteDirectory(instanceProfileDir);
			} catch (IOException ioException) {
				File oldProfileDir = new File(instanceProfileDir.getParentFile(),
						instanceProfileDir.getName() + ".old." + System.currentTimeMillis());
				if (instanceProfileDir.renameTo(oldProfileDir)) {
					logger.warning("could not delete profileDir: " + ioException.getMessage() + "; renamed it to "
							+ oldProfileDir);
				} else {
					logger.severe("could not delete profileDir: " + ioException.getMessage());
				}
			}
		}
	}

	private void addBasisAndUrePaths(ProcessBuilder processBuilder) throws IOException {
		// see
		// http://wiki.services.openoffice.org/wiki/ODF_Toolkit/Efforts/Three-Layer_OOo
		File basisLink = new File(officeHome, "basis-link");
		if (!basisLink.isFile()) {
			logger.fine(
					"no %OFFICE_HOME%/basis-link found; assuming it's OOo 2.x and we don't need to append URE and Basic paths");
			return;
		}
		String basisLinkText = FileUtils.readFileToString(basisLink).trim();
		File basisHome = new File(officeHome, basisLinkText);
		File basisProgram = new File(basisHome, "program");
		File ureLink = new File(basisHome, "ure-link");
		String ureLinkText = FileUtils.readFileToString(ureLink).trim();
		File ureHome = new File(basisHome, ureLinkText);
		File ureBin = new File(ureHome, "bin");
		Map<String, String> environment = processBuilder.environment();
		// Windows environment variables are case insensitive but Java maps are
		// not :-/
		// so let's make sure we modify the existing key
		String pathKey = "PATH";
		for (String key : environment.keySet()) {
			if ("PATH".equalsIgnoreCase(key)) {
				pathKey = key;
			}
		}
		String path = environment.get(pathKey) + ";" + ureBin.getAbsolutePath() + ";" + basisProgram.getAbsolutePath();
		logger.fine(String.format("setting %s to \"%s\"", pathKey, path));
		environment.put(pathKey, path);
	}

	public boolean isRunning() {
		if (process == null) {
			return false;
		}
		return getExitCode() == null;
	}

	private class ExitCodeRetryable extends Retryable {

		private int exitCode;

		protected void attempt() throws TemporaryException, Exception {
			try {
				exitCode = process.exitValue();
			} catch (IllegalThreadStateException illegalThreadStateException) {
				throw new TemporaryException(illegalThreadStateException);
			}
		}

		public int getExitCode() {
			return exitCode;
		}

	}

	public Integer getExitCode() {
		try {
			return process.exitValue();
		} catch (IllegalThreadStateException exception) {
			return null;
		}
	}

	public int getExitCode(long retryInterval, long retryTimeout) throws RetryTimeoutException {
		try {
			ExitCodeRetryable retryable = new ExitCodeRetryable();
			retryable.execute(retryInterval, retryTimeout);
			return retryable.getExitCode();
		} catch (RetryTimeoutException retryTimeoutException) {
			throw retryTimeoutException;
		} catch (Exception exception) {
			throw new OfficeException("could not get process exit code", exception);
		}
	}

	public long findOfficeProcessId() throws IOException {
		ProcessQuery processQuery = new ProcessQuery(executable.getName(), unoUrl.getAcceptString());
		return processManager.findPid(processQuery);
	}

	public int forciblyTerminate(long retryInterval, long retryTimeout) throws IOException, RetryTimeoutException {

		// Does the process still exist?
		ProcessQuery processQuery = new ProcessQuery(executable.getName(), unoUrl.getAcceptString());
		long pidFound = processManager.findPid(processQuery);

		if (pidFound == ProcessManager.PID_NOT_FOUND) {
			logger.severe("Process " + processQuery + " is no longer running");
			return 0;
		}

		if (pidFound != pid) {
			logger.severe("Looking for '" + processQuery + "' pid=" + pid + " but found pid=" + pidFound);
		}

		logger.info(String.format("trying to forcibly terminate process: '" + processQuery + "'"
				+ (pid != PID_UNKNOWN ? " (pid " + pid + ")" : "")));
		processManager.kill(process, pid);
		return getExitCode(retryInterval, retryTimeout);
	}

}
