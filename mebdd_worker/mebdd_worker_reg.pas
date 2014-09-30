unit mebdd_worker_reg;

(*
	MEB-DD's register-related functions.
*)

{$MODE OBJFPC}

interface

function read_registry_hklm_value_str (key_path:String; key_name:String):String;
function read_registry_hkcu_value_int (key_path:String; key_name:String):Integer;
function write_registry_hkcu_value_int (key_path:String; key_name:String; new_value:Integer):Boolean;

implementation

uses registry, Classes, SysUtils, Windows;

function is_windows_64: boolean;
  {
  Detect if we are running on 64 bit Windows or 32 bit Windows,
  independently of bitness of this program.
  Original source:
  http://www.delphipraxis.net/118485-ermitteln-ob-32-bit-oder-64-bit-betriebssystem.html
  modified for FreePascal in German Lazarus forum:
  http://www.lazarusforum.de/viewtopic.php?f=55&t=5287
  }
{$ifdef WIN32} //Modified KpjComp for 64bit compile mode
type
	TIsWow64Process = function( // Type of IsWow64Process API fn
		Handle: Windows.THandle; var Res: Windows.BOOL): Windows.BOOL; stdcall;
var
	IsWow64Result: Windows.BOOL; // Result from IsWow64Process
	IsWow64Process: TIsWow64Process; // IsWow64Process fn reference
begin
	// Try to load required function from kernel32
	IsWow64Process := TIsWow64Process(Windows.GetProcAddress(Windows.GetModuleHandle('kernel32'),'IsWow64Process'));
    if Assigned(IsWow64Process) then
		begin
			// Function is implemented: call it
			if not IsWow64Process(Windows.GetCurrentProcess, IsWow64Result) then
				raise SysUtils.Exception.Create('IsWindows64: bad process handle');
			// Return result of function
			Result := IsWow64Result;
		end
	else
		// Function not implemented: can't be running on Wow64
		Result := False;
{$else} //if were running 64bit code, OS must be 64bit :)
begin
	Result := True;
{$endif}
end;

function read_registry_hklm_value_str (key_path:String; key_name:String):String;

var
	key_value: string='';
	registry: TRegistry;
	
begin
	if is_windows_64() then
		registry := TRegistry.Create(KEY_READ OR KEY_WOW64_64KEY)
	else
		registry := TRegistry.Create(KEY_READ OR KEY_WOW64_32KEY);
		
	{$I-}
	try
		registry.RootKey := HKEY_LOCAL_MACHINE;
		
		if registry.OpenKeyReadOnly(key_path) then
			key_value := registry.ReadString(key_name);
	except
		On ERegistryException do
			key_value := '';
	end;

	registry.Free;
	
	read_registry_hklm_value_str := key_value;
end;

function read_registry_hkcu_value_int (key_path:String; key_name:String):Integer;

var
	key_value: integer=-1;
	registry: TRegistry;
	
begin
	if is_windows_64() Then
		registry := TRegistry.Create(KEY_READ OR KEY_WOW64_64KEY)
	else
		registry := TRegistry.Create(KEY_READ OR KEY_WOW64_32KEY);
	
	{$I-}
	try
		registry.RootKey := HKEY_CURRENT_USER;
		
		if registry.OpenKeyReadOnly(key_path) then
			key_value := registry.ReadInteger(key_name);
	except
		On ERegistryException do
			key_value := -1;
	end;
	
	registry.Free;
	
	read_registry_hkcu_value_int := key_value;
end;

function write_registry_hkcu_value_int (key_path:String; key_name:String; new_value:Integer):Boolean;

var
	write_ok: boolean=false;
	registry: TRegistry;
	
begin
	if is_windows_64() Then
		registry := TRegistry.Create(KEY_WRITE OR KEY_WOW64_64KEY)
	else
		registry := TRegistry.Create(KEY_WRITE OR KEY_WOW64_32KEY);

	try
		registry.RootKey := HKEY_CURRENT_USER;
		
		if registry.OpenKey(key_path, True) then
			begin
				registry.WriteInteger(key_name, new_value);
				registry.CloseKey;
				write_ok := true;
			end;
	except
		On ERegistryException do
			write_ok := False;
	end;
	
	registry.Free;
	
	write_registry_hkcu_value_int := write_ok;
end;

end.
