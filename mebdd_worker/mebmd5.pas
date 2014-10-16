Program mebmd5;

{$MODE OBJFPC}

Uses md5, SysUtils;

Function count_file_md5(filepath:AnsiString):AnsiString;

begin
	// If no file exist return ''
	If FileExists(filepath) Then
		begin
			count_file_md5 := LowerCase(MD5Print(MD5File(ParamStr(1))));
		end
	Else
		count_file_md5 := '';
	
end;

var
	md5_file, md5_param: AnsiString;
	
begin
	Case ParamCount of
		1: begin
			// One parameter: calculate MD5 and print
			md5_file := LowerCase(count_file_md5(ParamStr(1)));
			
			If md5_file = '' Then
				begin
					// Could not get MD5
					Writeln('-');
					Halt(0);
				end
			Else
				begin
					// We have the MD5
					Writeln(md5_file);
					Halt(1);
				end;
		end;
		2: begin
			// Two parameters: calculate and verify MD5
			md5_file := LowerCase(count_file_md5(ParamStr(1)));
			md5_param := LowerCase(ParamStr(2));
			
			If md5_file = '' Then
				begin
					// Could not get MD5
					writeln('md5 is empty for file '+md5_file);
					Halt(0);
				end
			Else
				begin
					// We have both MD5s
					If md5_file = md5_param Then
						begin
							// Sums are equal
							Halt(1);
						end
					Else
						begin
							// Sums are not equal
							Halt(2);
						end;
				end;
			end;
		end;

		// Zero or too many parameters
	Writeln('usage: mebmd5 filename [md5sum]');
	Writeln('');
	Writeln('Calculates MD5 sum of the file.');
	Writeln('* If no "md5sum" is given prints the calculated MD5 to');
	Writeln('  standard output and exists with exit code 1.');
	Writeln('* If "md5sum" is given calculates the MD5 and verifies');
	Writeln('  the calculated sum with the given. If the sums match');
	Writeln('  exits with code 1, otherwise 2.');
	Writeln('Exit code 0 signals an error.');
	
	Halt(0);
end.

