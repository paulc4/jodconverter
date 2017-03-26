# JODConverter

This is derived from JODConverter 3.0 beta, forked from https://github.com/mirkonasato/jodconverter.

## What is it?

JODConverter automates conversions between office document formats
using LibreOffice or Apache OpenOffice.  It runs up the Office application
in "headless" server mode and then makes requests for conversions by
sending requests oveer a socket.  The default port is 8100.

See http://jodconverter.googlecode.com for some documentation.


## Changes

The code has been updated as follows:

1. Compiles for Java 8
1. Recognises LibreOffice V4 and Open Office V4 when hunting for executables.
1. Fixed startup command in `OfficeProcess"
1. SIGAR libraries included (see below)
1. Wraps up the whole API into a convenient single class `OpenOfficeConverterUsingJOD3`

### SIGAR

One of the nice imprvements from JOD V2 is automated start and stop of the
server process. To do this it uses the SIGAR monitoring tool if available.
SIGAR in urn relies on Operating System facilities so it requires some
native libraries to be available.

The SIGAR native libraries/DLLs are included in a zip file in
`src/main/resources/sigar-native-lib.zip`.  At runtime, this zip
file is unpacked into `.jod-sigar-lib` in the user's home drectly.

Alternatively you can explicitly specify the location using 
`SigarProcessManager.setupLibraryPath(sigarNativeLibDir)`.

SIGAR (System Information Gatherer and Reporter) is a cross-platform,
cross-language library and command-line tool for accessing operating
system and hardware level information in Java, Perl and .NET.

Originally developed by Hyperic, the company was taken over by VMware.
It can be found here: https://sourceforge.net/projects/sigar.

## Licensing

JODConverter is open source software, you can redistribute it and/or
modify it under either (at your option) of the following licenses

1. The GNU Lesser General Public License v3 (or later)
   -> see LICENSE-LGPL.txt
2. The Apache License, Version 2.0
   -> see LICENSE-Apache.txt
