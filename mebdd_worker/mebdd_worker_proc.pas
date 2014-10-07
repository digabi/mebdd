unit mebdd_worker_proc;

(*
	MEB-DD's process related functions
*)

{$MODE OBJFPC}

interface

uses Classes;

const
	DRIVE_MAX = 99;
	
type
	Tdrive_string_arr = Array [0..DRIVE_MAX] of AnsiString;
	Tproc_output_arr = Array [0..DRIVE_MAX] of AnsiString;
	Texec_params = Record
		executable : AnsiString;
		parameters : TStringList;
	end;
	Texec_params_array = Array [0..DRIVE_MAX] of Texec_params;

var
	proc_log_file : AnsiString = '';
	
function is_pid_alive (const pid: DWORD; const proc_name:AnsiString): Boolean;
procedure proc_log_filename (filename:AnsiString);
procedure proc_log_truncate ();
procedure proc_log_writeln (message:AnsiString);
procedure proc_log_halt (e:Integer);
procedure proc_log_string (message:AnsiString);
function run_processes(proc_count:Word; cmdline:Texec_params_array):Tproc_output_arr;
function run_processes_count(proc_count:Word; cmdline:Texec_params_array; success_regexp:AnsiString; failure_regexp:AnsiString):Word;

implementation

uses Sysutils, Windows, JwaTlHelp32, Process, RegExpr;

function is_pid_alive (const pid: DWORD; const proc_name:AnsiString): Boolean;

var
	i: integer;
	_result: Boolean;
	CPID: DWORD;
	CProcName: array[0..259] of char;
	S: HANDLE;
	PE: TProcessEntry32;
  
begin
	_result := false;
	CProcName := '';
	S := CreateToolHelp32Snapshot(TH32CS_SNAPALL, 0); // Create snapshot
	PE.DWSize := SizeOf(PE); // Set size before use
	I := 1;
	if Process32First(S, PE) then
		repeat
			CProcName := PE.szExeFile;
			CPID := PE.th32ProcessID;

			if (CPID=pid) and (UpperCase(CProcName) = UpperCase(proc_name)) then
				_result := true;

			Inc(i);
		until not Process32Next(S, PE);
	CloseHandle(S);

	is_pid_alive := _result;
end;

procedure proc_log_filename (filename:AnsiString);

begin
	proc_log_file := filename;
end;


procedure proc_log_truncate ();

var
	f: TextFile;
	
begin
	if proc_log_file <> '' then
		begin
			AssignFile(f, proc_log_file);
			ReWrite(f);
			CloseFile(f);
		end;
end;


procedure proc_log_writeln (message:AnsiString);

begin
	writeln(message);
	proc_log_string(message);
end;

procedure proc_log_halt (e:Integer);

begin
	proc_log_string('Exiting with code #'+IntToStr(e));
	halt(e);
end;


procedure proc_log_string (message:AnsiString);

var
	f: TextFile;
	
begin
	if proc_log_file <> '' then
		begin
			AssignFile(f, proc_log_file);
			{I+}
			try
				Append(f);
			except
				on E: EInOutError do ReWrite(f);
			end;
			
			writeln(f, message);
			CloseFile(f);
		end;
end;


function run_processes(proc_count:Word; cmdline:Texec_params_array):Tproc_output_arr;

var
	proc_handle: Array [0..DRIVE_MAX] of TProcess;
	proc_output: Array [0..DRIVE_MAX] of TMemoryStream;
	proc_finished: Array [0..DRIVE_MAX] of Boolean;
	proc_output_bytes: Array [0..DRIVE_MAX] of LongInt;
	proc_output_bytes_available: LongInt;
	proc_output_bytes_read: LongInt;
	proc_output_stringlist: TStringList;
	proc_output_result: Tproc_output_arr;
	finished_processes:Word;
	n:Word;
	
	cycle_string: Array of AnsiString;
	cycle_n: Word;

	// Helper procedure to read process output from the process handle and add it to
	// the process output stream (proc_output)
	procedure add_proc_output_data(proc_index:Word);
	
	begin
		proc_output_bytes_available := proc_handle[proc_index].Output.NumBytesAvailable;
		if proc_output_bytes_available > 0 then
			begin
				// The stream has data

				// Allocate space
				proc_output[proc_index].SetSize(proc_output_bytes[proc_index] + proc_output_bytes_available);
				
				proc_output_bytes_read := proc_handle[proc_index].Output.Read((proc_output[proc_index].Memory + proc_output_bytes[proc_index])^, proc_output_bytes_available);
				
				if proc_output_bytes_read > 0 then
					proc_output_bytes[proc_index] := proc_output_bytes[proc_index] + proc_output_bytes_read;
			end;
	end;
	
begin
	// Debugging code
	proc_log_string('Running following processes:');
	for n := 0 to proc_count-1 do
		begin
			proc_log_string(cmdline[n].executable + ': ' + cmdline[n].parameters.CommaText);
		end;

	// Set cycle string array
	SetLength(cycle_string, 4);
	cycle_string[0] := '.';
	cycle_string[1] := 'o';
	cycle_string[2] := 'O';
	cycle_string[3] := 'o';
	cycle_n := 0;
	
	// Start processes
	for n := 0 to proc_count-1 do
		begin
			// Initialise process data
			proc_handle[n] := TProcess.Create(nil);
			proc_handle[n].Executable := cmdline[n].executable;
			proc_handle[n].Parameters := cmdline[n].parameters;
			
			// Put all I/O through pipes
			proc_handle[n].Options := [poUsePipes, poStdErrToOutPut];
			
			proc_handle[n].Execute;
			
			// Initialise output stream
			proc_output[n] := TMemoryStream.Create;
			proc_output_bytes[n] := 0;
			
			// Turn finish flag down
			proc_finished[n] := false;
		end;
	
	// Wait for all processed to end
	finished_processes := 0;
	
	while finished_processes < proc_count do
		begin
			// Count cycle pointer
			cycle_n := cycle_n + 1;
			cycle_n := cycle_n mod Length(cycle_string);

			for n := 0 to proc_count-1 do
				begin
					
					// Read output (if available)
					add_proc_output_data(n);
					
					if proc_handle[n].Running then
						begin
							// This process is still running
							//writeln(IntToStr(proc_handle[n].ProcessID) + ': ' + cycle_string[cycle_n]);
						end
					else
						begin
							// This process has finished
							
							// If this is the first time to encounter this, do things
							if not proc_finished[n] then
								begin
									Inc(finished_processes);
									
									// Read output (if available)
									add_proc_output_data(n);
									
									// Turn flag up so we don't run this again
									proc_finished[n] := True;
									
									proc_output_stringlist := TStringList.Create;
									proc_output_stringlist.LoadFromStream(proc_output[n]);
									
									// Add exit code to proc_output
									proc_output_stringlist.Add('exit_code:'+IntToStr(proc_handle[n].ExitStatus));
									
									// Log process output
									proc_log_string('Process finished:');
									proc_log_string(proc_output_stringlist.Text);
									
									// Store process output
									proc_output_result[n] := proc_output_stringlist.Text;
									
								end;
						end;
					
				end;
		
			// Update status line
			write('Jobs running: ' + IntToStr(proc_count - finished_processes) + ' ' + cycle_string[cycle_n] + AnsiString(#13));

			// Wait a bit until checking the process status again
			Sleep(1000);
		end;

	run_processes := proc_output_result;
end;

function run_processes_count(proc_count:Word; cmdline:Texec_params_array; success_regexp:AnsiString; failure_regexp:AnsiString):Word;
var
	process_output: Tproc_output_arr;
	n: Word;
	RegexObj: TRegExpr;
	process_failed: Boolean;
	failed_processes: Word;
	
begin
	// Call underlying process executer
	process_output := run_processes(proc_count, cmdline);
	
	// Go through all output
	failed_processes := 0;
	
	for n := 0 to proc_count-1 do
		begin
	
			// Check from output whether process output if process has failed
			process_failed := False;
			
			// First check the success regexp
			if success_regexp <> '' then
				begin
					RegexObj := TRegExpr.Create;
					RegexObj.Expression := success_regexp;
					RegexObj.ModifierI := True;
					if RegexObj.Exec(process_output[n]) then
						begin
							// This process reported success
							process_failed := False;
						end
					else
						begin
							// Oops, the output misses the success match
							process_failed := True;
						end;
				end;
				
			if failure_regexp <> '' then
				begin
					RegexObj := TRegExpr.Create;
					RegexObj.Expression := failure_regexp;
					RegexObj.ModifierI := True;
					if RegexObj.Exec(process_output[n]) then
						begin
							// This process reported failure
							process_failed := True;
						end;
				end;
			
			RegexObj.Free;
			
			if process_failed then
				begin
				proc_log_string('Process finish status: FAILED');
				Inc(failed_processes);
				end
			else
				begin
				proc_log_string('Process finish status: SUCCESS');
				end;
		end;

	run_processes_count := proc_count - failed_processes;
end;


end.
