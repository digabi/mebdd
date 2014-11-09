Program mebdd_read;

{$MODE OBJFPC}

Uses SysUtils, Classes, Windows, md5customised;


Function calculate_disk_md5 (disk_name:AnsiString; disk_bytes:Int64):AnsiString;

Var
	disk_handle: THandle;
	ctx: TMDContext;
	digest: TMD5Digest;
	error_code: Int64;
	bytes_read, bytes_read_total: Int64;
	buffer: array [0..65535] of Byte;   // 262144 (240Kb) or 131072 (120Kb buffer) or 65536 (64Kb buffer)

begin
	// Words with Windows 7 32 bit, Windows 8 64bit
	disk_handle := CreateFile(PChar(disk_name), GENERIC_READ,
		FILE_SHARE_READ, nil, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0);

	if disk_handle = INVALID_HANDLE_VALUE then
		begin
			// Error codes: http://msdn.microsoft.com/en-us/library/windows/desktop/ms681381(v=vs.85).aspx
			error_code := GetLastError();
			writeln('error opening disk "'+disk_name+'": #' + IntToStr(error_code));
			calculate_disk_md5 := '';
			Exit;
		end;
	
	// Make sure we're in the beginning of the file
	FileSeek(disk_handle, 0, 0);
	bytes_read_total := 0;
	
	// Initialise MD5 counter
	MD5Init(ctx);
	
	repeat
		if (disk_bytes - bytes_read_total) < SizeOf(buffer) then
			begin
				bytes_read := FileRead(disk_handle, buffer, (disk_bytes - bytes_read_total));  // This is the last buffer
			end
		else
			begin
				bytes_read := FileRead(disk_handle, buffer, SizeOf(buffer));  // There are still buffers to read
			end;

		if bytes_read = -1 then
			begin
				// Error while reading
				error_code := GetLastError();
				writeln('error while reading disk: #' + IntToStr(error_code));
				calculate_disk_md5 := '';
				Exit;
			end
		else
			// Step 2 : No read errors, so now we update the MD5 calculator
			begin
				inc(bytes_read_total, bytes_read);
				MD5Update(ctx, buffer, bytes_read);
			end;
		
	until (bytes_read_total = disk_bytes);

	if not disk_handle = INVALID_HANDLE_VALUE then CloseHandle(disk_handle);
	
	// Report success (the digest)
	MD5Final(ctx, digest);
	calculate_disk_md5 := Lowercase(MD5Print(digest));
end;

Procedure show_usage;

begin
	writeln('usage: mebdd_read source_disk_path source_disk_length md5_string');
	writeln('');
	writeln('mebdd_read dumps "source_disk_length" bytes from "source_disk_path"');
	writeln('(e.g. \\.\PHYSICALDRIVE1) and calculates the MD5 of the data.');
	writeln('If the calculated MD5 equals to "md5_string" returns SUCCESS.');
	writeln('');
	writeln('exit codes:');
	writeln('0 - failed (probably an unhandled exception)');
	writeln('1 - success');
	writeln('255 - failed (error reported to standard output');
	Writeln('');
	Writeln('mebdd_read build ' + {$I %DATE%} + ' ' + {$I %TIME%});
	
	Halt(255);
end;


Var
	disk_path: AnsiString;
	disk_bytes: Int64;
	digest_expected: AnsiString;
	digest_disk: AnsiString;
	
begin
	// Read command line parameters
	if ParamCount <> 3 then
		show_usage;
	
	// Show build datetime to help debugging
	Writeln('mebdd_read build ' + {$I %DATE%} + ' ' + {$I %TIME%});

	disk_path := ParamStr(1);
	disk_bytes := 0;	// Set a default value
	digest_expected := Lowercase(ParamStr(3));
	
	// Get bytes from the command line
	if not TryStrToInt64(ParamStr(2), disk_bytes) then
		begin
			writeln('Error: ' + ParamStr(2) + ' is not a number');
			writeln('');
			
			show_usage;
		end;
	
	// Call the digest calculator
	digest_disk := calculate_disk_md5(disk_path, disk_bytes);
	if (digest_disk = digest_expected) then
		begin
			writeln('Verification ok');
			Halt(1);
		end
	else
		begin
			writeln('Verification failed (calculated digest: ' + digest_disk);
			Halt(255);
		end;
end.
