# mebdd

VBscript based frontend for dd, curl to easily create DigabiOS boot memory sticks.

## Structure

mebdd is build to two parts. `mebdd\mebdd.hta` is a Windows HTA script written in VBscript. It creates the GUI that calls other helper binaries. `mebdd_worker\mebdd_worker.pas` is a helper program that controls the disk processes.

mebdd functions and their implementation:

* When the user wants to download the image the `mebdd.hta` calls cURL binary to retrieve the disk image from a pre-defined location. The calling script checks the exit code of cURL and reports this to the user.
* When writing the image the `mebdd.hta` calls `mebdd_worker.exe` with image and USB disk paths (e.g. \\.\PHYSICALDISK1). The `mebdd_worker`
	* asks the user (UAC) for Adminisrative privileges in order to successfully execute `DISKPART` and `dd`
	* uses `DISKPART.EXE` (the standard Windows command-line tool for manipulating storage devices) to check that the disk ID belongs to an USB drive
	* cleans the existing MBR and partition table with CLEAN command
	* writes the image to the device using `dd.exe`
	* follows the parallel subprocesses of the activities
* The disk clean function is equal to image writing. A 512 bytes long file of zeros (`zero_image.iso`) is written to disk.

## Compiling mebdd_worker.exe

mebdd_worker is written in [Lazarus](http://lazarus.freepascal.org/), a Free Pascal based Delphi compatible IDE. It probably compiles with standard Free Pascal but I haven't tried this. The development is done with v2.6.4. The script `mebdd_worker\compile.bat` feeds the needed parameters to the compiler and linker.

## Packaging

The project can be packaged for delivery using Inno Setup (v5.5). Please note the control file `install\mebdd_X.iss`.

To create a package you need following external binaries:
* `curl.exe` from [cURL project](http://curl.haxx.se/)
* `dd.exe` from [chrysocome.net](http://www.chrysocome.net/dd)
* `libeay32.dll` and `ssleay32.dll` (used by cURL) from [OpenSSL project](http://www.openssl.org/related/binaries.html)

