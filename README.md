# mebdd

VBscript based frontend for `dd`, `unzip` and `curl` to easily create DigabiOS boot memory sticks.

## Structure

mebdd is build to two parts. `mebdd\mebdd.hta` is a Windows HTA script written in VBscript. It creates the GUI that calls other helper utilities. These binaries are written in Free Pascal. Currently there are three helper utilities:
* `mebdd_worker.exe` writes images (by calling `dd`), verifies them (`mebdd_read`) and creates FAT filesystems (by calling Windows `DISKPART`)
* `mebdd_read.exe` verifies written image: reads disk images, calculates their MD5 sum and verifies the calculated MD5 to the given MD5 calculated from the image file
* `mebmd5.exe` calculates the MD5 sum of a given file

mebdd functions and their implementation:

* When the user wants to download the image the `mebdd.hta` calls cURL binary to retrieve the disk image from a pre-defined location. The calling script checks the exit code of cURL and reports this to the user.
* When writing the image the `mebdd.hta` calls `mebdd_worker.exe` with image and USB disk paths (e.g. \\.\PHYSICALDISK1). The `mebdd_worker`
	* asks the user (UAC) for Adminisrative privileges in order to successfully execute `DISKPART` and `dd`
	* uses `DISKPART.EXE` (the standard Windows command-line tool for manipulating storage devices) to check that the disk ID belongs to an USB drive
	* cleans the existing MBR and partition table with CLEAN command
	* writes the image to the device using `dd.exe`
	* verifies the written image using `mebdd_read.exe`
	* follows the parallel subprocesses of the activities
* The disk clean function calls `DISKPART` and feeds it a script created just before execution. The script removes all existing partitions and creates a partition with a maximum size. Finally it creates a filesystem with a workstation's default filesystem type (typically FAT).

## Compiling helper utilities

mebdd_worker is written in [Lazarus](http://lazarus.freepascal.org/), a Free Pascal based Delphi compatible IDE. It probably compiles with standard Free Pascal but I haven't tried this. The development is done with v2.6.4.

As the building procedure is quite straightforward there is no Makefile but each utility has its own compile script.

* To compile `mebdd_worker` execute `compile.bat`
* To compile `mebdd_read` execute `compile_mebdd_read.bat`
* To compile `mebmd5` execute `compile_md5.bat`

## Packaging

The project can be packaged for delivery using Inno Setup (v5.5). Please note the control file `install\mebdd_X.iss`.

To create a package you need following external binaries:
* `curl.exe` from [cURL project](http://curl.haxx.se/)
* `dd.exe` from [chrysocome.net](http://www.chrysocome.net/dd)
* `libeay32.dll` and `ssleay32.dll` (used by cURL) from [OpenSSL project](http://www.openssl.org/related/binaries.html)

## Configuring mebdd.hta

`mebdd.hta` can be configured using the mebdd.ini file which is located in the install directory (currently \Program Files\MEB Finland\mebdd).

```
# Lines starting with # and ; are comments
; This is another comment

[*System]
	ImageUpdateMethod = individual 
	; "group" = All updated or missing images must be updated
	; "individual" = You can update individual images

; This defines an image. Image tag (e.g. "ktp") is used internally to distinct
; images from each other.
[ktp]
	Legend = KTP
	; Legend: name of the image
	Description = Koetilan palvelin
	; Description: A slightly longer description of the image
	URLimage = http://kurko.digabi.fi/latest/ktp.zip
	; URLimage: Defines an URL of the image. If image ends with .zip
	;           mebdd unzips a file defined with "LocalFile" from this zip
	URLMD5 = http://kurko.digabi.fi/latest/ktp.zip.md5
	; URLMD5: Defines an URL of the MD5 file. The file should contain
	;         only one MD5 sum
	LocalFile = ktp.dd
	; LocalFile: Local file name. The image files are stored in
	;            AppData\Local\MEB Finland\mebdd. This also specifies the
	;            image filename to unzip from the zipped image package.
	RegRemoteMD5 = MD5_digabiktp_remote
	; RegRemoteMD5: MD5 value of the latest known remote image retrieved
	;               from "URLMD5".
	;              "RegRemoteMD5" specifies a registry key name containing
	;              the MD5 sum. The keys are stored in
	;              HKEY_CURRENT_USER\Software\MEB Finland\mebdd
	RegLocalMD5 = MD5_digabiktp_local
	; RegLocalMD5: MD5 value of the locally present image file. This MD5
	;              belongs to a downloaded image file. If the image file
	;              was a zip file the unzipped image file typically has
	;              different MD5 sum. This is not stored to any registry
	;              value.
	;              "RegLocalMD5" specifies a registry key name containing
	;              the MD5 sum.

[koe]
	; This defines another image
	Legend = KOE
	Description = Kokelaan päätelaite
	URLimage = http://kurko.digabi.fi/latest/koe.zip
	URLMD5 = http://kurko.digabi.fi/latest/koe.zip.md5
	LocalFile = koe.dd
	RegRemoteMD5 = MD5_digabikoe_remote
	RegLocalMD5 = MD5_digabikoe_local

[digabi64]
	; Gosh! Third image!
	Legend = DigabiOS 64 bit
	Description = DigabiOS demo, 64 bit
	URLimage = https://digabi.fi/64bit
	URLMD5 = https://digabi.fi/64bit/md5
	LocalFile = DigabiOS-64bit.dd
	RegRemoteMD5 = MD5_digabi64_remote
	RegLocalMD5 = MD5_digabi64_local

[digabi32]
	Legend = DigabiOS 32 bit
	Description = DigabiOS demo, 32 bit
	URLimage = https://digabi.fi/32bit
	URLMD5 = https://digabi.fi/32bit/md5
	LocalFile = DigabiOS-32bit.dd
	RegRemoteMD5 = MD5_digabi32_remote
	RegLocalMD5 = MD5_digabi32_local
```

## Debugging

`mebdd.hta` and `mebdd_worker.exe` write a log file to `%TEMP%\mebdd.txt`.

