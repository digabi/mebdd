Program mebdd_read;

{$MODE OBJFPC}

Uses SysUtils, Classes, Windows;


Function copy_disk_to_file (disk_name:AnsiString; disk_bytes:Int64; file_name: AnsiString):Boolean;

Var
	disk_name_wc: Array [0..MAX_PATH] of WideChar;
	disk_handle: THandle;
	file_handle: TFileStream;
	error_code: Int64;
	bytes_read, bytes_read_total: Int64;
	buffer: array [0..65535] of Byte;   // 262144 (240Kb) or 131072 (120Kb buffer) or 65536 (64Kb buffer)

begin
	StringToWideChar(disk_name, disk_name_wc, MAX_PATH);
	
	disk_handle := CreateFileW(PWideChar(disk_name_wc), FILE_READ_DATA,
		FILE_SHARE_READ, nil, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0);

	if disk_handle = INVALID_HANDLE_VALUE then
		begin
			// Error codes: http://msdn.microsoft.com/en-us/library/windows/desktop/ms681381(v=vs.85).aspx
			error_code := GetLastError();
			writeln('error opening disk "'+disk_name+'": #' + IntToStr(error_code));
			copy_disk_to_file := False;
			Exit;
		end;
	
	Try
		file_handle := TFileStream.Create(file_name, fmCreate);
	Except
		on EFOpenError do
			begin
				writeln('Could not create file ' + file_name);
				copy_disk_to_file := False;
				Exit;
			end;
		else
			begin
				writeln('Unknown error while creating file ' + file_name);
				copy_disk_to_file := False;
				Exit;
			end;
	End;
	
	// Make sure we're in the beginning of the file
	FileSeek(disk_handle, 0, 0);
	bytes_read_total := 0;
	
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
				copy_disk_to_file := False;
				Exit;
			end
		else
			// Step 2 : No read errors, so now we write data to file...
			begin
				try
					file_handle.WriteBuffer(buffer, SizeOf(buffer));
				except
					on EStreamError do
						begin
							writeln('Error while writing to file ' + file_name);
							copy_disk_to_file := False;
							Exit;
						end
					else
						begin
							writeln('Unknown error while writing to file ' + file_name);
							copy_disk_to_file := False;
							Exit;
						end;
				end;
				
				inc(bytes_read_total, bytes_read);
			end;
		
	until (bytes_read_total = disk_bytes);

	if not disk_handle = INVALID_HANDLE_VALUE then CloseHandle(disk_handle);
	file_handle.Free;
	
	// Report success
	copy_disk_to_file := True;
end;

Procedure show_usage;

begin
	writeln('usage: mebdd_read source_disk_path source_disk_length destination_file_path');
	writeln('');
	writeln('mebdd_read dumps "source_disk_length" bytes from "source_disk_path"');
	writeln('(e.g. \\.\PHYSICALDRIVE1) to regular file "destination_file_path".');
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
	disk_bytes: LongInt;
	file_path: AnsiString;
	
begin
	// Read command line parameters
	if ParamCount <> 3 then
		show_usage;
	
	// Show build datetime to help debugging
	Writeln('mebdd_read build ' + {$I %DATE%} + ' ' + {$I %TIME%});

	disk_path := ParamStr(1);
	disk_bytes := 0;	// Set a default value
	file_path := ParamStr(3);
	
	// Get bytes from the command line
	if not TryStrToInt(ParamStr(2), disk_bytes) then
		begin
			writeln('Error: ' + ParamStr(2) + ' is not a number');
			writeln('');
			
			show_usage;
		end;
	
	// Call the copier function
	if (copy_disk_to_file(disk_path, disk_bytes, file_path)) then
		begin
			writeln('Dump ok');
			Halt(1);
		end
	else
		begin
			writeln('Dump failed');
			Halt(255);
		end;
end.
