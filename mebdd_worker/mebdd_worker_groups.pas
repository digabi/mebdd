unit mebdd_worker_groups;

(*

Taken from http://stackoverflow.com/questions/8288525/how-to-get-windows-user-privileges-information-with-lazarus-free-pascal

*)

{$MODE OBJFPC}

interface

uses Windows;

const
	SECURITY_NT_AUTHORITY: TSIDIdentifierAuthority = (Value: (0, 0, 0, 0, 0, 5));
	SECURITY_BUILTIN_DOMAIN_RID = $00000020;
	DOMAIN_ALIAS_RID_ADMINS     = $00000220;
	DOMAIN_ALIAS_RID_USERS      = $00000221;
	DOMAIN_ALIAS_RID_GUESTS     = $00000222;
	DOMAIN_ALIAS_RID_POWER_USERS= $00000223;

function  UserInGroup(Group :DWORD) : Boolean;

implementation

uses SysUtils, Classes;


function CheckTokenMembership(TokenHandle: THandle; SidToCheck: PSID; var IsMember: BOOL): BOOL; stdcall; external advapi32;

function  UserInGroup(Group :DWORD) : Boolean;

var
	pIdentifierAuthority :TSIDIdentifierAuthority;
	pSid : Windows.PSID;
	IsMember    : BOOL;

begin
	pIdentifierAuthority := SECURITY_NT_AUTHORITY;
	Result := AllocateAndInitializeSid(pIdentifierAuthority,2, SECURITY_BUILTIN_DOMAIN_RID, Group, 0, 0, 0, 0, 0, 0, pSid);
	try
		if Result then
			if not CheckTokenMembership(0, pSid, IsMember) then //passing 0 means which the function will be use the token of the calling thread.
			Result:= False
		else
			Result:=IsMember;
	finally
		FreeSid(pSid);
	end;
end;

end.
