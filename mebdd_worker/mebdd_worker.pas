program mebdd_worker;

{$MODE OBJFPC}
{$R manifest.rc}
{$R mebdd_worker.rc}

uses Sysutils, Classes, RegExpr, mebdd_worker_reg, mebdd_worker_proc, mebdd_worker_groups;

const
	RETRIES = 3;
	
var
	path_mebdd: AnsiString='';
	path_diskpart: AnsiString='';
	path_cmd: AnsiString='';
	path_image: AnsiString='';
	path_drives: Tdrive_string_arr;
	path_drive_count: Word=0;
	switch_verify: Boolean=False;
	

function file_is_readable(file_path:AnsiString):boolean;
var
	_result:boolean=true;
	f:TextFile;
	buffer:AnsiString;
	
begin
	{$I+}
	try
		AssignFile(f, file_path);
		Reset(f);
		read(f, buffer);
		CloseFile(f);
	except
		on EInOutError do
			_result := false;
	end;
	
	file_is_readable := _result;
end;

function get_file_size(file_path:AnsiString):LongInt;
var
	_result:LongInt;
	f:File of Byte;

begin	
	{$I+}
	try
		Assign(f,file_path);
		Reset(f);
		_result := FileSize(f);
		Close(f);
	except
		// FIXME
	end;
	
	get_file_size := _result;
end;

function is_worker_already_running():Boolean;

var
	registry_pid:Integer;
	_result:Boolean;
	
begin
	_result := false;
	registry_pid := read_registry_hkcu_value_int('Software\MEB Finland\mebdd', 'running_pid');
	
	if registry_pid <= 0 then
		begin
			// Key was not found
			_result := false;
		end
	else
		begin
			if is_pid_alive(registry_pid, 'iexplore.exe') then
				_result := true;
		end;
		
	is_worker_already_running := _result;
end;

function devicepath_to_disknumber(devicepath:AnsiString):AnsiString;
var
	RegexObj: TRegExpr;
	str_devicepath_disknumber: AnsiString;
	_result: AnsiString;
	
begin
	// Default result 'not found'
	_result := '';
	
	RegexObj := TRegExpr.Create;
	RegexObj.Expression := '^\\\\\.\\PHYSICALDRIVE(\d)$';
	RegexObj.ModifierI := True;
	if RegexObj.Exec(devicepath) then
		begin
			str_devicepath_disknumber := RegexObj.Match[1];
			If str_devicepath_disknumber <> '' Then
				_result := str_devicepath_disknumber;
		end;

	devicepath_to_disknumber := _result;
end;

function drive_paths_are_usb(drive_count:Word; drives:Tdrive_string_arr):Boolean;
var
	cmdline:Texec_params_array;
	diskpart_script:AnsiString;
	parameters:TStringList;
	process_output:Tproc_output_arr;
	n, match_count:Word;
	f: TextFile;
	RegexObj: TRegExpr;
	
begin
	diskpart_script := IncludeTrailingPathDelimiter(GetEnvironmentVariable('TEMP'))+'mebdd_worker_'+IntToStr(GetProcessID)+'.txt';
	
	// Open DISKPART script file
	AssignFile(f, diskpart_script);
	ReWrite(f);

	for n:=0 to drive_count-1 do
		begin
			// Create DISKPART script for this drive
			
			If devicepath_to_disknumber(drives[n]) = '' Then
				begin
					proc_log_writeln('Device path "'+drives[n]+'" does not resolve to a disk number');
					drive_paths_are_usb := False;
					Exit;
				end;
			
			writeln(f, 'SELECT DISK '+devicepath_to_disknumber(drives[n]));
			writeln(f, 'DETAIL DISK');
		end;
		
	writeln(f, 'EXIT');
	CloseFile(f);
	
	// Set process executable + parameters
	cmdline[0].executable := path_diskpart;
	parameters := TStringList.Create;
	parameters.Add('/s');
	parameters.Add(diskpart_script);
	cmdline[0].parameters := parameters;
	
	// Run processes
	process_output := run_processes(1, cmdline);
	
	// Delete DISKPART script
	If not DeleteFile(diskpart_script) Then
		proc_log_string('Warning: failed to delete temporary file "'+diskpart_script+'"');

	// Check process output
	RegexObj := TRegExpr.Create;
	RegexObj.Expression := ': USB\r';
	if RegexObj.Exec(process_output[0]) then
		begin
			// We have at least one match
			match_count := 1;
			
			while RegexObj.ExecNext do
				Inc(match_count);
			
			if match_count = drive_count then
				drive_paths_are_usb := True
			else
				drive_paths_are_usb := False;
		end;

	RegexObj.Free;
end;

function get_file_digest(path_image:AnsiString):AnsiString;
var
	image_details, image_details_this: AnsiString;
	image_filedate: LongInt;
	digest: AnsiString;
	cmdline:Texec_params_array;
	parameters:TStringList;
	process_output:Tproc_output_arr;
	RegexObj: TRegExpr;
	
begin
	// Try to read an existing digest
	image_details := read_registry_hkcu_value_string('Software\MEB Finland\mebdd', 'image_details');
	digest := read_registry_hkcu_value_string('Software\MEB Finland\mebdd', 'image_md5');
	
	image_filedate := FileAge(path_image);
	
	if image_filedate = -1 then
		image_details_this := path_image + ':' + IntToStr(get_file_size(path_image))
	else
		image_details_this := path_image + ':' + IntToStr(get_file_size(path_image)) + ':' + DateTimeToStr(FileDateTodateTime(image_filedate));
	
	if image_details <> image_details_this then
		begin
		// The digest does not belong to this file
		digest := '';
		end;
		
	// If we did not find a digest, calculate it
	if digest = '' then
		begin
			// Run MD5 as an external process
			cmdline[0].executable := IncludeTrailingPathDelimiter(path_mebdd)+'mebmd5.exe';
			parameters := TStringList.Create;
			parameters.Add(path_image);
			cmdline[0].parameters := parameters;
			
			process_output := run_processes(1, cmdline);
			
			// Search the MD5 digest from the output
			RegexObj := TRegExpr.Create;
			RegexObj.Expression := '^([0123456789abcdef]+)';
			RegexObj.ModifierI := True;
			if RegexObj.Exec(process_output[0]) then
				begin
					digest := RegexObj.Match[1];
					
					// Store digest to registry
					if not write_registry_hkcu_value_string('Software\MEB Finland\mebdd', 'image_details', image_details_this) then
						proc_log_writeln('Warning: Could not store digest to registry - image_details');
					if not write_registry_hkcu_value_string('Software\MEB Finland\mebdd', 'image_md5', digest) then
						proc_log_writeln('Warning: Could not store digest to registry - image_md5');
				end;
		end;

	get_file_digest := digest;
end;


	
function clear_partition_tables(drive_count:Word; drives:Tdrive_string_arr):Boolean;
var
	cmdline:Texec_params_array;
	diskpart_script:AnsiString;
	parameters:TStringList;
	finished_processes:Word;
	n:Word;
	f: TextFile;
	retry_counter: Integer;
	
begin
	diskpart_script := IncludeTrailingPathDelimiter(GetEnvironmentVariable('TEMP'))+'mebdd_worker_'+IntToStr(GetProcessID)+'.txt';
	
	// Open DISKPART script file
	AssignFile(f, diskpart_script);
	ReWrite(f);

	for n:=0 to drive_count-1 do
		begin
			// Create DISKPART script for this drive
			
			If devicepath_to_disknumber(drives[n]) = '' Then
				begin
					proc_log_writeln('Device path "'+drives[n]+'" does not resolve to a disk number');
					clear_partition_tables := False;
					Exit;
				end;
			
			writeln(f, 'SELECT DISK '+devicepath_to_disknumber(drives[n]));
			writeln(f, 'CLEAN');
		end;
		
	writeln(f, 'EXIT');
	CloseFile(f);
	
	// Set process executable + parameters
	cmdline[0].executable := path_diskpart;
	parameters := TStringList.Create;
	parameters.Add('/s');
	parameters.Add(diskpart_script);
	cmdline[0].parameters := parameters;
	
	// Run processes in a repeat loop
	retry_counter := 0;
	
	repeat
		// Increase retry counter + write to log
		Inc(retry_counter);
		proc_log_string('Retry '+IntToStr(retry_counter)+':');
		
		// Run process
		finished_processes := run_processes_count(1, cmdline, 'exit_code:0', '');
	
		if finished_processes = 1 then
			clear_partition_tables := True
		else
			clear_partition_tables := False;

	until ((retry_counter > RETRIES) or clear_partition_tables);
	
	if (not clear_partition_tables) then
		proc_log_string('All retries failed.');

	// Delete DISKPART script
	If not DeleteFile(diskpart_script) Then
		proc_log_string('Warning: failed to delete temporary file "'+diskpart_script+'"');

end;

function write_disk_image(drive_count:Word; drives:Tdrive_string_arr; image_file:AnsiString):Boolean;
var
	dd_cmdline:Texec_params_array;
	parameters:TStringList;
	finished_processes:Word;
	n:Word;
	retry_counter: Integer;
	
begin
	for n:=0 to drive_count-1 do
		begin
			dd_cmdline[n].executable := IncludeTrailingPathDelimiter(path_mebdd)+'dd.exe';
			parameters := TStringList.Create;
			parameters.Add('if="'+image_file+'"');
			parameters.Add('of="'+drives[n]+'"');
			parameters.Add('bs=1M');
			dd_cmdline[n].parameters := parameters;
		end;
	

	// Run processes in a repeat loop
	retry_counter := 0;
	
	repeat
		// Increase retry counter + write to log
		Inc(retry_counter);
		proc_log_string('Retry '+IntToStr(retry_counter)+':');
		
		// Run processes
		finished_processes := run_processes_count(drive_count, dd_cmdline, '', '.*Error.*');
	
		if finished_processes = drive_count then
			write_disk_image := True
		else
			write_disk_image := False;
			
	until ((retry_counter > RETRIES) or write_disk_image);

	if (not write_disk_image) then
		proc_log_string('All retries failed.');
end;

function verify_disk_image(drive_count:Word; drives:Tdrive_string_arr; image_path: AnsiString; digest:AnsiString):Boolean;
var
	dd_cmdline:Texec_params_array;
	parameters:TStringList;
	finished_processes:Word;
	temporary_filename: Tdrive_string_arr;
	n:Word;
	file_size: LongInt;
	retry_counter: Integer;
	
begin
	// Calculate image size in order to know how many bytes we read from the USB
	file_size := get_file_size(image_path);

	for n:=0 to drive_count-1 do
		begin
			// Create temporary name for this image
			
			temporary_filename[n] := IncludeTrailingPathDelimiter(GetEnvironmentVariable('TEMP'))+'mebdd_worker_image_'+IntToStr(GetProcessID)+'_'+IntToStr(n)+'.img';
			
			dd_cmdline[n].executable := path_cmd;
			parameters := TStringList.Create;
			parameters.Add('/c');
			parameters.Add('"');
			parameters.Add(IncludeTrailingPathDelimiter(path_mebdd)+'dd.exe');
			parameters.Add('if="'+drives[n]+'"');
			parameters.Add('of="'+temporary_filename[n]+'"');
			parameters.Add('bs='+IntToStr(file_size));
			parameters.Add('count=1');
			parameters.Add('&');
			parameters.Add(IncludeTrailingPathDelimiter(path_mebdd)+'mebmd5.exe');
			parameters.Add(temporary_filename[n]);
			parameters.Add(digest);
			parameters.Add('"');
			dd_cmdline[n].parameters := parameters;
		end;
	
	// Run processes in a repeat loop
	retry_counter := 0;
	
	repeat
		// Increase retry counter + write to log
		Inc(retry_counter);
		proc_log_string('Retry '+IntToStr(retry_counter)+':');
		
		// Run processes
		finished_processes := run_processes_count(drive_count, dd_cmdline, 'exit_code:1', '');
	
		if finished_processes = drive_count then
			verify_disk_image := True
		else
			verify_disk_image := False;
	until ((retry_counter > RETRIES) or verify_disk_image);

	if (not verify_disk_image) then
		proc_log_string('All retries failed.');
	
	// Delete temporary files
	for n:=0 to drive_count-1 do
		begin
			// Delete temporary file
			if (not DeleteFile(temporary_filename[n])) then
				proc_log_string('Warning: Could not delete temporary image file '+temporary_filename[n]);
		end;
end;


var
	n:Word;
	param_n:Word;
	image_digest:AnsiString;
	
begin
	// Init log file
	proc_log_filename(IncludeTrailingPathDelimiter(GetEnvironmentVariable('TEMP'))+'mebdd_worker.txt');
	proc_log_string('Execution started');
	
	// Check if user is admin
	if not UserInGroup(DOMAIN_ALIAS_RID_ADMINS) then
		begin
			proc_log_writeln('Error: You have to have administrative privileges to run this program');
			proc_log_halt(1);
		end;
		
	// Read paths from the registry
	path_mebdd := read_registry_hklm_value_str('SOFTWARE\MEB Finland\mebdd', 'path_mebdd');
	path_diskpart := read_registry_hklm_value_str('SOFTWARE\MEB Finland\mebdd', 'path_diskpart');
	path_cmd := read_registry_hklm_value_str('SOFTWARE\MEB Finland\mebdd', 'path_cmd');
	
	if path_mebdd = '' then
		begin
			proc_log_writeln('Error: MEB-DD installation path is not set');
			proc_log_halt(2);
		end;

	if path_diskpart = '' then
		begin
			proc_log_writeln('Error: DISKPART path is not set');
			proc_log_halt(3);
		end;
	
	if path_cmd = '' then
		begin
			proc_log_writeln('Error: CMD path is not set');
			proc_log_halt(3);
		end;
	
	if is_worker_already_running() then
		begin
			proc_log_writeln('Error: mebdd_worker is already running');
			proc_log_halt(4);
		end;
		
	// Read command line parameters
	if (ParamCount = 0) then
		begin
			writeln('usage: mebdd_worker [-v] image_to_write.iso device_to_write1 [device_to_writeN...]');
			proc_log_halt(11);
		end;
	
	param_n := 1;
	
	// Do we have -v = verify?
	if ParamStr(param_n) = '-v' then
		begin
			switch_verify := True;
			param_n := 2;
		end
	else
		switch_verify := False;
		
	path_image := ParamStr(param_n);
	Inc(param_n);

	for n:=param_n to ParamCount do
		begin
			path_drives[path_drive_count] := ParamStr(n);
			path_drive_count := path_drive_count+1;
		end;
		
	if path_image = '' then
		begin
			proc_log_writeln('Error: image path missing');
			proc_log_halt(5);
		end;
	if not file_is_readable(path_image) then
		begin
			proc_log_writeln('Error: image file "'+path_image+'" is not readable');
			proc_log_halt(6);
		end;
	
	// Set lock
	if not write_registry_hkcu_value_int('Software\MEB Finland\mebdd', 'running_pid', GetProcessID) then
		begin
			proc_log_writeln('Error: PID write failed');
			proc_log_halt(7);
		end;
	
	// Make sure that path_drives resolve to USB memory sticks
	proc_log_writeln('Checking drive path(s):');
	if drive_paths_are_usb(path_drive_count, path_drives) then
		begin
			proc_log_writeln('Success!                                 ');
		end
	else
		begin
			proc_log_writeln('Error: Drive path check failed');
			proc_log_halt(8);
		end;
		
	// Count/create file MD5 digest
	if (switch_verify) then
		begin
			proc_log_writeln('Getting image file digest:');
			image_digest := get_file_digest(path_image);
			if (image_digest = '') then
				begin
					proc_log_writeln('Error: Could not get image digest');
					proc_log_halt(13);
				end
			else
				proc_log_writeln('Success!                                 ');
		end
	else
		image_digest := '';
	
	// Clear partition tables
	proc_log_writeln('Clearing master boot record(s) and partition table(s):');
	if clear_partition_tables(path_drive_count, path_drives) then
		begin
			proc_log_writeln('Success!                                 ');
			// Sleep 3 seconds to wait all pending disk activity to end
			Sleep(3000);
		end
	else
		begin
			proc_log_writeln('Error: Failed to clear partition table');
			proc_log_halt(9);
		end;
	
	// Write image
	proc_log_writeln('Writing disk image(s):');
	if write_disk_image(path_drive_count, path_drives, path_image) then
		begin
			proc_log_writeln('Success!                                 ');
		end
	else
		begin
			proc_log_writeln('Error: Failed to write disk image');
			proc_log_halt(10);
		end;
	
	// Verify image
	if (switch_verify) then
		if image_digest = '' then
			begin
				proc_log_writeln('Error: Image verify failed as MD5 was not calculated');
				proc_log_halt(12);
			end
		else
			begin
				proc_log_writeln('Verifying disk image(s):');
				if verify_disk_image(path_drive_count, path_drives, path_image, image_digest) then
					begin
						proc_log_writeln('Success!                                 ');
					end
				else
					begin
						proc_log_writeln('Error: disk image verification failed');
						proc_log_halt(13);
					end;
			end;

	// Clear lock
	write_registry_hkcu_value_int('Software\MEB Finland\mebdd', 'running_pid', 0);
	
	proc_log_string('Normal termination');
	proc_log_halt(255);
end.
