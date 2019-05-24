
#include <iostream>
#include <iomanip>
#include <string>
#include <algorithm>
#include <fstream>
#include <Shlobj.h>
#include <Windows.h>
#include <stdlib.h>
#include <dos.h>
#include <VersionHelpers.h>
#include <conio.h>
#include <direct.h>

#include <odbcinst.h>
#include <sql.h>
#include <sqlext.h>
#include <sqltypes.h>
#include <sqlucode.h>
#include <msdasql.h>
#include <msdadc.h>

using namespace std;

#ifndef UNICODE  
typedef std::string String;
#else
typedef std::wstring String;
#endif

//System Variables
TCHAR wWinDir[255]; String WinDir;
TCHAR wDataDir[255]; String DataDir;
TCHAR wPgmsDir[255]; String PgmsDir;

//ADOConnection Variables
TCHAR wConnectThru[255]; String ConnectThru;
TCHAR wDatabaseName[255]; String DatabaseName;
TCHAR wDataSource[255]; String DataSource;

//INSTALL Variables
TCHAR wSQLbuild[255]; String SQLbuild;
TCHAR wInstalledVersion[255]; String InstalledVersion;
TCHAR wBuild[255]; String Build;
TCHAR wServicePack[255]; String ServicePack;

//Windows Specific Variables
String rUsername; //USERNAME
String rUserprofile; //USERPROFILE
String rLocalAppData; //LOCALAPPDATA
String rHostname; //COMPUTERNAME
String rSystemRoot; //SYSTEMROOT 

bool debugOn = false;
String runningVersion = "v0.0.52";

bool run = true;

void menu();
void SQLmenu();
void header();
void cls();
void pause();
void adminPriv();
void winSvrChk();
int getSysInfo();
string curDir();
void resizeWindow();
void exit();
void logOutput();

int getOmate32();
int OMRunAsAdmin();
int duplicateINI();
void delTmpFiles();
void runBatchFiles();
void enableLinkedConnections();

void showSQLErrorMsg(unsigned int handleType, const SQLHANDLE& handle);
void OMUserLoginTest();
void clearLogonInfo();
void closeAllOpenExams();

/*
  FEATURES & NOTES TO ADD:

- set folder permissions for PgmsDir
- set file permissions for Omate32.ini in WinDir
*/


int main()
{
	resizeWindow();
	winSvrChk();
	adminPriv();
	getSysInfo();

	while (run == true) {
		cls();
		header();
		getOmate32();
		logOutput();
		menu();
	}
	exit();
	return 0;
}

void header() {
	if (debugOn == true) {
		cout << "Eyefinity Officemate Suite Information and Diagnostics Tool " << runningVersion << " (Debug ON)" << endl;
		cout << "--------------------------------------------------------------------------------" << string(1, '\n');
	}
	else {
		cout << "Eyefinity Officemate Suite Information and Diagnostics Tool " << runningVersion << endl;
		cout << "--------------------------------------------------------------------------------" << string(1, '\n');
	}
}

void cls() {

	CHAR fill = ' ';
	COORD tl = { 0,0 };
	CONSOLE_SCREEN_BUFFER_INFO s;
	HANDLE console = GetStdHandle(STD_OUTPUT_HANDLE);

	GetConsoleScreenBufferInfo(console, &s);
	DWORD written, cells = s.dwSize.X * s.dwSize.Y;
	FillConsoleOutputCharacter(console, fill, cells, tl, &written);
	FillConsoleOutputAttribute(console, s.wAttributes, cells, tl, &written);
	SetConsoleCursorPosition(console, tl);
}

void pause() {
	char a;
	cout << string(2, '\n') << "Press any Key to Continue... ";
	a = _getch();


}

void resizeWindow() {
	HWND console = GetConsoleWindow();
	RECT r;
	GetWindowRect(console, &r);

	//MoveWindow(window_handle, x, y, width, height, redraw_window);
	MoveWindow(console, r.left, r.top, 800, 550, TRUE);
}

void adminPriv() {
	if (IsUserAnAdmin() == false) {
		header();
		cout << String(2, '\n') << " CANNOT CONTINUE: RUN AS ADMINISTRATOR.";
		pause();
		exit(5);
	}
}

void winSvrChk() {

	if (IsWindowsServer() == true) {
		header();
		cout << String(2, '\n') << " CANNOT CONTINUE: NOT DESIGNED FOR WINDOWS SERVER.";
		pause();
		exit();
	}
}

void menu() {
	int x = 8;
	cout << String(2, '\n');
	cout << setw(8) << left << "Option" << setw(15) << left << "Solutions";
	cout << endl << "--------------------------------------------------------------------------------";
	cout << endl << setw(x) << left << " A." << setw(40) << left << "Create SetRunAsAdmin.reg to set all main OM executables to run as Admin";
	cout << endl << setw(x) << left << " B." << setw(40) << left << "Create EnableLinkedConnections.reg to merge mapped drive permissions";
	cout << endl << setw(x) << left << " C." << setw(40) << left << "Check for Duplicate Omate32.ini in common folders";
	cout << endl << setw(x) << left << " D." << setw(40) << left << "Delete .tmp files on root of C:\\ left from running Reports";
	cout << endl << setw(x) << left << " E." << setw(40) << left << "Run all ~Reg*.bat files located in Officemate Directory (v12 and Below)";
	cout << endl << setw(x) << left << " F." << setw(40) << left << "Test OM_USER and SA connection and login to OfficeMate SQL Database";
	cout << endl;
	cout << endl << setw(x) << left << " S." << setw(40) << left << "Enter Menu for SQL Queries and Fixes";
	cout << endl;
	cout << endl << setw(x) << left << " Z." << setw(40) << left << "Exit" << endl << endl;

	char menuopt;
	char confirmopt;

	cout << setw(10) << left << "Enter menu option from above to run solution: ";
	cin >> menuopt;
	menuopt = toupper(menuopt);
	cout << endl;

	if (menuopt != 'Z') {
		cout << setw(10) << left << "Are you sure you want to run solution - " << menuopt << "? Y or N:";
		cin >> confirmopt; cout << endl;
		confirmopt = toupper(confirmopt);
		if (confirmopt != 'Y') {
			main();;
		}
	}

	switch (menuopt)
	{
	case 'A':
		OMRunAsAdmin();
		break;
	case 'B':
		enableLinkedConnections();
		break;
	case 'C':
		duplicateINI();
		break;
	case 'D':
		delTmpFiles();
		break;
	case 'E':
		runBatchFiles();
		break;
	case 'F':
		OMUserLoginTest();
		break;
	case 'S':
		SQLmenu();
		break;
	case 'Z':
		exit();
		break;
	}
}

void SQLmenu() {
	cls();
	header();

	int x = 8;
	cout << left << setw(10) << "SQL Queries and Fixes Menu" << String(3,'\n'); 
	cout << setw(8) << left << "Option" << setw(15) << left << "Solutions";
	cout << endl << "--------------------------------------------------------------------------------";
	cout << endl << setw(x) << left << " A." << setw(40) << left << "Clear User Seat Count (ClearLogonInfo)";
	cout << endl << setw(x) << left << " B." << setw(40) << left << "Close all open exams in ExamWriter";
	cout << endl;
	cout << endl << setw(x) << left << " Z." << setw(40) << left << "Exit" << endl << endl;

	char menuopt;
	char confirmopt;

	cout << setw(10) << left << "Enter menu option from above to run solution: ";
	cin >> menuopt;
	menuopt = toupper(menuopt);

	if (menuopt != 'Z') {
		cout << setw(10) << left << "Are you sure you want to run solution - " << menuopt << "? Y or N:";
		cin >> confirmopt; cout << endl;
		confirmopt = toupper(confirmopt);
		if (confirmopt != 'Y') {
			SQLmenu();;
		}
	}

	switch (menuopt) {
	case 'A':
		clearLogonInfo();
		break;
	case 'B':
		closeAllOpenExams();
		break;
	case 'Z':
		main();
		break;
	}

	cout << String(2, '\n');
}

int getSysInfo() {

	size_t requiredSize;

	//Get rUsername 
	char* username;
	getenv_s(&requiredSize, NULL, 0, "username");
	if (requiredSize == 0) { cout << setw(10) << left << "ERROR: Could not get USERNAME from environment"; }
	username = (char*)malloc(requiredSize * sizeof(char));
	if (!username) { cout << setw(10) << left << "ERROR: Failed to allocate memory!\n"; }
	getenv_s(&requiredSize, username, requiredSize, "username");
	rUsername = username;
	//cout << setw(10) << left << "Username variable is: " << username << endl;
	free(username); requiredSize = 0;

	//Get rUserprofile 
	char* userprofile;
	getenv_s(&requiredSize, NULL, 0, "userprofile");
	if (requiredSize == 0) { cout << setw(10) << left << "ERROR: Could not get USERPROFILE from environment"; }
	userprofile = (char*)malloc(requiredSize * sizeof(char));
	if (!userprofile) { cout << setw(10) << left << "ERROR: Failed to allocate memory!\n"; }
	getenv_s(&requiredSize, userprofile, requiredSize, "userprofile");
	rUserprofile = userprofile;
	//cout << setw(10) << left << "Userprofile variable is: " << userprofile << endl;
	free(userprofile); requiredSize = 0;

	//Get rLocalAppdata 
	char* localappdata;
	getenv_s(&requiredSize, NULL, 0, "localappdata");
	if (requiredSize == 0) { cout << setw(10) << left << "ERROR: Could not get LocalAppData from environment"; }
	localappdata = (char*)malloc(requiredSize * sizeof(char));
	if (!localappdata) { cout << setw(10) << left << "ERROR: Failed to allocate memory!\n"; }
	getenv_s(&requiredSize, localappdata, requiredSize, "localappdata");
	rLocalAppData = localappdata;
	//cout << setw(10) << left << "Localappdata variable is: " << localappdata << endl;
	free(localappdata); requiredSize = 0;

	//Get rHostname 
	char* hostname;
	getenv_s(&requiredSize, NULL, 0, "computername");
	if (requiredSize == 0) { cout << setw(10) << left << "ERROR: Could not get Hostname from environment"; }
	hostname = (char*)malloc(requiredSize * sizeof(char));
	if (!hostname) { cout << setw(10) << left << "ERROR: Failed to allocate memory!\n"; }
	getenv_s(&requiredSize, hostname, requiredSize, "computername");
	rHostname = hostname;
	//cout << setw(10) << left << "Hostname variable is: " << hostname << endl;
	free(hostname); requiredSize = 0;

	//Get rSystemRoot 
	char* systemroot;
	getenv_s(&requiredSize, NULL, 0, "systemroot");
	if (requiredSize == 0) { cout << setw(10) << left << "ERROR: Could not get SystemRoot from environment"; }
	systemroot = (char*)malloc(requiredSize * sizeof(char));
	if (!systemroot) { cout << setw(10) << left << "ERROR: Failed to allocate memory!\n"; }
	getenv_s(&requiredSize, hostname, requiredSize, "systemroot");
	rSystemRoot = systemroot;
	//cout << setw(10) << left << "SystemRoot variable is: " << hostname << endl;
	free(systemroot); requiredSize = 0;

	return 0;
}

string curDir() {
	char buffer[MAX_PATH];
	GetModuleFileName(NULL, buffer, MAX_PATH);
	string::size_type pos = string(buffer).find_last_of("\\/");
	return string(buffer).substr(0, pos);
}

void exit() {
	cout << string(1, '\n') << "Press any Key to Exit... ";
	char a;
	a = _getch();
	exit(EXIT_SUCCESS);
}

void logOutput() {
	//Output info             1

	cout << setw(20) << left << "WinDir           = " << left << WinDir << endl;
	cout << setw(20) << left << "DataDir          = " << left << DataDir << endl;
	cout << setw(20) << left << "PgmsDir          = " << left << PgmsDir << endl << endl;
	//cout << setw(20) << left << "ConnectThru      = " << left << ConnectThru << endl;
	cout << setw(20) << left << "DatabaseName     = " << left << DatabaseName << endl;
	cout << setw(20) << left << "DataSource       = " << left << DataSource << endl;
	cout << setw(20) << left << "SQL_Build        = " << left << SQLbuild << endl << endl;
	cout << setw(20) << left << "InstalledVersion = " << left << InstalledVersion << endl;
	cout << setw(20) << left << "Build            = " << left << Build << endl;
	cout << setw(20) << left << "ServicePack      = " << left << ServicePack << endl;

}

int getOmate32() {
	cout << "Gathering Information from C:\\Windows\\Omate32.ini..." << string(2, '\n');

	//System Section
	GetPrivateProfileString(TEXT("System"), TEXT("WinDir"), TEXT("INI NOT FOUND"), wWinDir, 255, TEXT("Omate32.ini"));
	WinDir = wWinDir;
	GetPrivateProfileString(TEXT("System"), TEXT("DataDir"), TEXT("INI NOT FOUND"), wDataDir, 255, TEXT("Omate32.ini"));
	DataDir = wDataDir;
	GetPrivateProfileString(TEXT("System"), TEXT("PgmsDir"), TEXT("INI NOT FOUND"), wPgmsDir, 255, TEXT("Omate32.ini"));
	PgmsDir = wPgmsDir;

	//ADOConnection Section
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("ConnectThru"), TEXT("INI NOT FOUND"), wConnectThru, 255, TEXT("Omate32.ini"));
	ConnectThru = wConnectThru;
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("DatabaseName"), TEXT("INI NOT FOUND"), wDatabaseName, 255, TEXT("Omate32.ini"));
	DatabaseName = wDatabaseName;
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("DataSource"), TEXT("INI NOT FOUND"), wDataSource, 255, TEXT("Omate32.ini"));
	DataSource = wDataSource;

	//Install Section
	GetPrivateProfileString(TEXT("Install"), TEXT("SQL_Build"), TEXT("INI NOT FOUND"), wSQLbuild, 255, TEXT("Omate32.ini"));
	SQLbuild = wSQLbuild;
	GetPrivateProfileString(TEXT("Install"), TEXT("InstalledVersion"), TEXT("INI NOT FOUND"), wInstalledVersion, 255, TEXT("Omate32.ini"));
	InstalledVersion = wInstalledVersion;
	GetPrivateProfileString(TEXT("Install"), TEXT("Build"), TEXT("INI NOT FOUND"), wBuild, 255, TEXT("Omate32.ini"));
	Build = wBuild;
	GetPrivateProfileString(TEXT("Install"), TEXT("Service Pack"), TEXT("INI NOT FOUND"), wServicePack, 255, TEXT("Omate32.ini"));
	ServicePack = wServicePack;

	transform(WinDir.begin(), WinDir.end(), WinDir.begin(), (int(*)(int))toupper);
	transform(DataDir.begin(), DataDir.end(), DataDir.begin(), (int(*)(int))toupper);
	transform(PgmsDir.begin(), PgmsDir.end(), PgmsDir.begin(), (int(*)(int))toupper);

	transform(ConnectThru.begin(), ConnectThru.end(), ConnectThru.begin(), (int(*)(int))toupper);
	transform(DatabaseName.begin(), DatabaseName.end(), DatabaseName.begin(), (int(*)(int))toupper);
	transform(DataSource.begin(), DataSource.end(), DataSource.begin(), (int(*)(int))toupper);

	transform(SQLbuild.begin(), SQLbuild.end(), SQLbuild.begin(), (int(*)(int))toupper);
	transform(InstalledVersion.begin(), InstalledVersion.end(), InstalledVersion.begin(), (int(*)(int))toupper);
	transform(ServicePack.begin(), ServicePack.end(), ServicePack.begin(), (int(*)(int))toupper);

	if ((DataDir == "INI NOT FOUND") || (PgmsDir == "INI NOT FOUND")) {
		cout << setw(10) << left << "Omate32.ini not found in C:\\Windows. Please correct Omate32.ini to proceed" << endl;
		cout << setw(10) << left << "Is Officemate\\ExamWriter installed on this PC? Are you cloud hosted?";
		pause();
		exit(1);
	}

	return 0;
}

int OMRunAsAdmin() {

	string regPgmsDir = PgmsDir;
	int strLength = regPgmsDir.length();

	int x; int pos = 0;
	size_t pos1 = 0, pos2 = 0;
	for (x = 0; x < strLength; x++) {
		pos1 = 0;
		pos1 = regPgmsDir.find("\\", pos, 1);
		pos2 = regPgmsDir.find("\\\\", pos, 2);
		if (pos1 != pos2) {
			pos = pos1 + 2;
			regPgmsDir.replace(pos1, 1, "\\\\");
		}
	}

	fstream SetRunAsAdmin;
	string filenamepath = curDir() + "\\SetRunAsAdmin.reg";
	SetRunAsAdmin.open(filenamepath, fstream::out);
	SetRunAsAdmin << "Windows Registry Editor Version 5.00\n";
	SetRunAsAdmin << " \n";
	SetRunAsAdmin << "[HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers]\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\examwriter.exe\"=\"RUNASADMIN\"\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\homeoffice.exe\"=\"RUNASADMIN\"\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\login.exe\"=\"RUNASADMIN\"\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\logindotnet.exe\"=\"RUNASADMIN\"\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\omate.exe\"=\"RUNASADMIN\"\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\query.exe\"=\"RUNASADMIN\"\n";

	SetRunAsAdmin.close();

	filenamepath = PgmsDir + "\\SetRunAsAdmin.reg";
	SetRunAsAdmin.open(filenamepath, fstream::out);
	SetRunAsAdmin << "Windows Registry Editor Version 5.00\n";
	SetRunAsAdmin << " \n";
	SetRunAsAdmin << "[HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers]\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\examwriter.exe\"=\"RUNASADMIN\"\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\homeoffice.exe\"=\"RUNASADMIN\"\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\login.exe\"=\"RUNASADMIN\"\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\logindotnet.exe\"=\"RUNASADMIN\"\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\omate.exe\"=\"RUNASADMIN\"\n";
	SetRunAsAdmin << "\"" + regPgmsDir + "\\\\query.exe\"=\"RUNASADMIN\"\n";

	cls();
	header();
	cout << endl;
	cout << setw(10) << left << " Successfully created SetRunAsAdmin.reg..." << string(2, '\n');
	cout << setw(10) << left << " Manually run SetRunAsAdmin.reg from " << PgmsDir << endl;
	cout << string(2, '\n');
	pause();

	SetRunAsAdmin.close();
	return 0;
}

int duplicateINI() {
	cls();
	header();
	cout << endl;

	ifstream vs;
	ifstream up;
	int fileCount = 0;

	string tmp = "\\virtualstore\\Windows\\Omate32.ini";
	string vspath = rLocalAppData + tmp;
	vs.open(vspath, ios::in);
	if (!vs.is_open()) {/*Code goes here that is executed if file note found in LocalAppData\virtualstore\Windows folder*/ }
	else {
		cout << setw(10) << left << " WARNING: Duplicate Omate32.ini found in %LocalAppData%\\Local\\VirtualStore\\Windows" << endl;
		fileCount++;
	}
	vs.close();

	tmp = "\\Windows\\Omate32.ini";
	string uppath = rUserprofile + tmp;
	up.open(uppath, ios::in);
	if (!up.is_open()) {/*Code goes here that is executed if file note found in Userprofile\Windows folder*/}
	else {
		cout << setw(10) << left << " WARNING: Duplicate Omate32.ini found in %userprofile%\\Windows Folder" << endl;
		fileCount++;
	}
	up.close();

	cout << string(2, '\n') << setw(10) << right << " Total Duplicate Omate32.ini Files Found: " << fileCount;
	pause();
	return 0;
}

void delTmpFiles()
{
	cls();
	header();
	cout << endl;

	const char *command1 = "@echo off && cd /d C:\\ && del *.tmp";
	cout << endl << setw(20) << left << "Deleting .tmp files, Please wait..." << string(2, '\n');
	system(command1);
	cout << endl << setw(20) << left << "Finished Deleting .tmp files on C:\\" << endl;
	pause();
}

void runBatchFiles() {
	//Runs all ~Reg batch files located in PgmsDir

	cls();
	header();
	cout << endl;
	cout << setw(10) << left << " Running all ~Reg .bat files in Officemate Directory" << string(2, '\n');

	// ~RegAllOME.NetDLLs.bat
	string batch1 = PgmsDir + "\\~RegAllOME.NetDLLs.bat";
	if ((int)ShellExecuteA(NULL, "open", batch1.c_str(), NULL, NULL, SW_NORMAL) > 32) {
		cout << setw(10) << left << " Running " << batch1 << endl;
	}

	// ~RegAllOME_DLLs.bat
	string batch2 = PgmsDir + "\\~RegAllOME_DLLs.bat";
	if ((int)ShellExecuteA(NULL, "open", batch2.c_str(), NULL, NULL, SW_NORMAL) > 32) {
		cout << setw(10) << left << " Running " << batch2 << endl;
	}

	// ~RegOME_DL_SysDLLs.bat
	string batch3 = PgmsDir + "\\~RegOME_DL_SysDLLs.bat";
	if ((int)ShellExecuteA(NULL, "open", batch3.c_str(), NULL, NULL, SW_NORMAL) > 32) {
		cout << setw(10) << left << " Running " << batch3 << endl;
	}

	// ~RegOME_DLLs4Svr.bat
	string batch4 = PgmsDir + "\\~RegOME_DLLs4Svr.bat";
	if ((int)ShellExecuteA(NULL, "open", batch4.c_str(), NULL, NULL, SW_NORMAL) > 32) {
		cout << setw(10) << left << " Running " << batch4 << endl;
	}

	// ~RegWinDLLs.bat
	string batch5 = PgmsDir + "\\~RegWinDLLs.bat";
	if ((int)ShellExecuteA(NULL, "open", batch5.c_str(), NULL, NULL, SW_NORMAL) > 32) {
		cout << setw(10) << left << " Running " << batch5 << endl;
	}

	string batch6 = PgmsDir + "\\~OM_REG.bat";
	if ((int)ShellExecuteA(NULL, "open", batch6.c_str(), NULL, NULL, SW_NORMAL) > 32) {
		cout << setw(10) << left << " Running " << batch6 << endl;
	}

	string batch7 = PgmsDir + "\\~OMRegAll.bat";
	if ((int)ShellExecuteA(NULL, "open", batch7.c_str(), NULL, NULL, SW_NORMAL) > 32) {
		cout << setw(10) << left << " Running " << batch7 << endl;
	}

	cout << endl;
	pause();
}

void enableLinkedConnections() {
	/* Exports EnableLinkedConnections.reg

	Example file:
	Windows Registry Editor Version 5.00

	[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System]
	"EnableLinkedConnections"=dword:00000001
	*/

	fstream enableLinkedConnectionsReg;
	string filenamepath = curDir() + "\\EnableLinkedConnections.reg";
	enableLinkedConnectionsReg.open(filenamepath, fstream::out);
	enableLinkedConnectionsReg << "Windows Registry Editor Version 5.00\n";
	enableLinkedConnectionsReg << " \n";
	enableLinkedConnectionsReg << "[HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System]\n\n";
	enableLinkedConnectionsReg << "\"EnableLinkedConnections\" = dword:00000001";

	enableLinkedConnectionsReg.close();

	filenamepath = PgmsDir + "\\EnableLinkedConnections.reg";
	enableLinkedConnectionsReg.open(filenamepath, fstream::out);
	enableLinkedConnectionsReg << "Windows Registry Editor Version 5.00\n";
	enableLinkedConnectionsReg << "\n";
	enableLinkedConnectionsReg << "[HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System]\n\n";
	enableLinkedConnectionsReg << "\"EnableLinkedConnections\"=dword:00000001";

	enableLinkedConnectionsReg.close();

	cls();
	header();
	cout << endl;
	cout << setw(10) << left << " Successfully created EnableLinkedConnections.reg..." << string(2, '\n');
	cout << setw(10) << left << " Manually run EnableLinkedConnections.reg from " << PgmsDir << endl;
	cout << string(2, '\n');
	pause();
}

void showSQLErrorMsg(unsigned int handleType, const SQLHANDLE& handle) {
	//Function shows error message and state derived from preliminary function return message
	SQLCHAR SQLState[1024];
	SQLCHAR message[1024];

	if (SQL_SUCCESS == SQLGetDiagRec(handleType, handle, 1, SQLState, NULL, message, 1024, NULL)) {
		cout << left << setw(10) << "ODBC Driver Message: " << message << string(2, '\n');;
		cout << left << setw(10) << "ODBC SQL State: " << SQLState << string(2, '\n');;
	}
}

void OMUserLoginTest() {
	// Connect to SQL database as OM_USER and then disconnect, and use SQL_ERROR info to show information
	cls();
	header();

	SQLHANDLE SQLEnvHandle = NULL;
	SQLHANDLE SQLConnectionHandle = NULL;
	SQLCHAR retConString[1024];

	string ConnectionString1 = "DRIVER={SQL Server}; SERVER=" + DataSource + "; DATABASE=" + DatabaseName + ";Uid=OM_USER;Pwd=OMSQL@2004;";
	string ConnectionString2 = "DRIVER={SQL Server}; SERVER=" + DataSource + "; DATABASE=" + DatabaseName + ";Uid=SA;Pwd=OMateSQL@2007;";

	if (SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &SQLEnvHandle) == SQL_ERROR) {
	}
	if (SQLSetEnvAttr(SQLEnvHandle, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, 0) == SQL_ERROR) {
	}
	if (SQLAllocHandle(SQL_HANDLE_DBC, SQLEnvHandle, &SQLConnectionHandle) == SQL_ERROR) {
	}
	if (SQLSetConnectAttr(SQLConnectionHandle, SQL_LOGIN_TIMEOUT, (SQLPOINTER)5, 0) == SQL_ERROR) {
	}

	cout << left << setw(40) << " Running OM_USER connection and login test: ";
	
	switch (SQLDriverConnect(SQLConnectionHandle, NULL, (SQLCHAR*)ConnectionString1.c_str(), SQL_NTS, retConString, 1024, NULL, SQL_DRIVER_NOPROMPT)) {
	case SQL_SUCCESS:
		cout << setw(35) << right << " Success - SQL_SUCESS" << string(2, '\n');;
		break;
	case SQL_SUCCESS_WITH_INFO:
		cout << setw(35) << right << " Success - SQL_SUCCESS_WITH_INFO" << string(2, '\n');;
		break;
	case SQL_NO_DATA_FOUND:
		cout << setw(35) << right << " Error - SQL_NO_DATA_FOUND" << string(2, '\n');;
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	case SQL_INVALID_HANDLE:
		cout << setw(35) << right << " Error - SQL_INVALID_HANDLE" << string(2, '\n');;
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	case SQL_ERROR:
		cout << setw(35) << right << " Error - SQL_ERROR" << string(2, '\n');
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	}

	SQLDisconnect(SQLConnectionHandle);
	SQLFreeHandle(SQL_HANDLE_DBC, SQLConnectionHandle);
	SQLFreeHandle(SQL_HANDLE_ENV, SQLEnvHandle);

	cout << endl << left << setw(40) << " Running SA connection and login test: ";

	if (SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &SQLEnvHandle) == SQL_ERROR) {
	}
	if (SQLSetEnvAttr(SQLEnvHandle, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, 0) == SQL_ERROR) {
	}
	if (SQLAllocHandle(SQL_HANDLE_DBC, SQLEnvHandle, &SQLConnectionHandle) == SQL_ERROR) {
	}
	if (SQLSetConnectAttr(SQLConnectionHandle, SQL_LOGIN_TIMEOUT, (SQLPOINTER)5, 0) == SQL_ERROR) {
	}

	switch (SQLDriverConnect(SQLConnectionHandle, NULL, (SQLCHAR*)ConnectionString1.c_str(), SQL_NTS, retConString, 1024, NULL, SQL_DRIVER_NOPROMPT)) {
	case SQL_SUCCESS:
		cout << setw(39) << right << " Success - SQL_SUCESS" << string(2, '\n');;
		break;
	case SQL_SUCCESS_WITH_INFO:
		cout << setw(39) << right << " Success - SQL_SUCCESS_WITH_INFO" << string(2, '\n');;
		break;
	case SQL_NO_DATA_FOUND:
		cout << setw(35) << right << " Error - SQL_NO_DATA_FOUND" << string(2, '\n');;
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	case SQL_INVALID_HANDLE:
		cout << setw(35) << right << " Error - SQL_INVALID_HANDLE" << string(2, '\n');;
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	case SQL_ERROR:
		cout << setw(35) << right << " Error - SQL_ERROR" << string(2, '\n');
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	}

	pause();
}

void clearLogonInfo() {
	cls();
	header();

	cout << "Datasource: " << DataSource << "           Database: " << DatabaseName << string(2, '\n');
	cout << "Displaying table dbo.user_logon_info: ";

	SQLHANDLE SQLEnvHandle = NULL;
	SQLHANDLE SQLConnectionHandle = NULL;
	SQLHANDLE SQLStatementHandle = NULL;
	SQLCHAR retConString[1024];

	string ConnectionString = "DRIVER={SQL Server}; SERVER=" + DataSource + "; DATABASE=" + DatabaseName + ";Uid=OM_USER;Pwd=OMSQL@2004;";

	if (SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &SQLEnvHandle) == SQL_ERROR) {
	}
	if (SQLSetEnvAttr(SQLEnvHandle, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, 0) == SQL_ERROR) {
	}
	if (SQLAllocHandle(SQL_HANDLE_DBC, SQLEnvHandle, &SQLConnectionHandle) == SQL_ERROR) {
	}
	if (SQLSetConnectAttr(SQLConnectionHandle, SQL_LOGIN_TIMEOUT, (SQLPOINTER)5, 0) == SQL_ERROR) {
	}

	switch (SQLDriverConnect(SQLConnectionHandle, NULL, (SQLCHAR*)ConnectionString.c_str(), SQL_NTS, retConString, 1024, NULL, SQL_DRIVER_NOPROMPT)) {
	case SQL_SUCCESS:
		cout << "Success - SQL_SUCESS" << string(2, '\n');;
		break;
	case SQL_SUCCESS_WITH_INFO:
		cout << "Success - SQL_SUCCESS_WITH_INFO" << string(2, '\n');;
		break;
	case SQL_NO_DATA_FOUND:
		cout << "Error - SQL_NO_DATA_FOUND" << string(2, '\n');;
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	case SQL_INVALID_HANDLE:
		cout << "Error - SQL_INVALID_HANDLE" << string(2, '\n');;
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	case SQL_ERROR:
		cout << "Error - SQL_ERROR" << string(2, '\n');
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	}

	if (SQLAllocHandle(SQL_HANDLE_STMT, SQLConnectionHandle, &SQLStatementHandle) == SQL_ERROR) {
		cout << left << setw(10) << "Error allocating SQL Statement Handle" << endl;
	}

	char SQLQuery[] = "SELECT logon_index, logon_machinename, convert(nvarchar(25),logon_datetime), App_Type FROM dbo.user_logon_info";

	// Display table and logins
	if (SQLExecDirect(SQLStatementHandle, (SQLCHAR*)SQLQuery, SQL_NTS) == SQL_ERROR) {
		cout << left << setw(10) << "Error executing SQL Statement Handle" << endl;
		showSQLErrorMsg(SQL_HANDLE_STMT, SQLStatementHandle);
	}
	else
	{
		int logon_index;
		char logon_machinename[25];
		char logon_datetime[25];
		char App_Type[25];

		cout << setw(15) << left << "logon_index" << setw(25) << "logon_machinename" << setw(25) << left << "logon_datetime" << setw(6) << left << "App_Type" << string(2, '\n');
		while (SQLFetch(SQLStatementHandle) == SQL_SUCCESS) {
			SQLGetData(SQLStatementHandle, 1, SQL_C_DEFAULT, &logon_index, sizeof(logon_index) + 1, NULL);
			SQLGetData(SQLStatementHandle, 2, SQL_C_DEFAULT, &logon_machinename, size(logon_machinename), NULL);
			SQLGetData(SQLStatementHandle, 3, SQL_C_DEFAULT, &logon_datetime, size(logon_datetime), NULL);
			SQLGetData(SQLStatementHandle, 4, SQL_C_DEFAULT, &App_Type, sizeof(App_Type), NULL);
			cout << setw(15) << left << logon_index << setw(25) << logon_machinename << setw(25) << left << logon_datetime << setw(6) << left << App_Type << endl;
		}
	}
	SQLFreeHandle(SQL_HANDLE_STMT, SQLStatementHandle);

	char clearOpt = 'N';
	cout << string(2, '\n');
	cout << " Enter Y to clear Users Logged In, Z to exit: ";
	cin >> clearOpt;
	clearOpt = toupper(clearOpt);

	if (clearOpt == 'Y') {
		// Truncate table
		char SQLQuery2[] = "TRUNCATE TABLE dbo.user_logon_info";

		cout << endl;
		cout << left << setw(10) << " Truncating table dbo.user_logon_info: ";

		if (SQLAllocHandle(SQL_HANDLE_STMT, SQLConnectionHandle, &SQLStatementHandle) == SQL_ERROR) {
			cout << left << setw(10) << "Error allocating SQL Statement Handle" << endl;
		}

		if (SQLExecDirect(SQLStatementHandle, (SQLCHAR*)SQLQuery2, SQL_NTS) == SQL_ERROR) {
			cout << left << setw(10) << "Error executing SQL Statement Handle" << endl;
			showSQLErrorMsg(SQL_HANDLE_STMT, SQLStatementHandle);
		}
		else {
			cout << "Success - SQL_SUCESS" << endl;
			clearLogonInfo();
		}
	}

	if (clearOpt != 'Z' && clearOpt != 'Y') {
		cout << endl << " Invalid character entered. Try again." << endl;
		pause();
		clearLogonInfo();
	}

	SQLFreeHandle(SQL_HANDLE_STMT, SQLStatementHandle);
	SQLDisconnect(SQLConnectionHandle);
	SQLFreeHandle(SQL_HANDLE_DBC, SQLConnectionHandle);
	SQLFreeHandle(SQL_HANDLE_ENV, SQLEnvHandle);

	if (clearOpt == 'Z') {
		SQLmenu();
	}
}

void closeAllOpenExams() {
	cls();
	header();

	cout << "Datasource: " << DataSource << "           Database: " << DatabaseName << string(2, '\n');
	cout << "Connecting to SQL Database: ";

	SQLHANDLE SQLEnvHandle = NULL;
	SQLHANDLE SQLConnectionHandle = NULL;
	SQLHANDLE SQLStatementHandle = NULL;
	SQLCHAR retConString[1024];
	int count = 0;

	string ConnectionString = "DRIVER={SQL Server}; SERVER=" + DataSource + "; DATABASE=" + DatabaseName + ";Uid=OM_USER;Pwd=OMSQL@2004;";

	if (SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &SQLEnvHandle) == SQL_ERROR) {
	}
	if (SQLSetEnvAttr(SQLEnvHandle, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, 0) == SQL_ERROR) {
	}
	if (SQLAllocHandle(SQL_HANDLE_DBC, SQLEnvHandle, &SQLConnectionHandle) == SQL_ERROR) {
	}
	if (SQLSetConnectAttr(SQLConnectionHandle, SQL_LOGIN_TIMEOUT, (SQLPOINTER)5, 0) == SQL_ERROR) {
	}

	switch (SQLDriverConnect(SQLConnectionHandle, NULL, (SQLCHAR*)ConnectionString.c_str(), SQL_NTS, retConString, 1024, NULL, SQL_DRIVER_NOPROMPT)) {
	case SQL_SUCCESS:
		cout << "Success - SQL_SUCESS" << string(2, '\n');;
		break;
	case SQL_SUCCESS_WITH_INFO:
		cout << "Success - SQL_SUCCESS_WITH_INFO" << string(2, '\n');;
		break;
	case SQL_NO_DATA_FOUND:
		cout << "Error - SQL_NO_DATA_FOUND" << string(2, '\n');;
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	case SQL_INVALID_HANDLE:
		cout << "Error - SQL_INVALID_HANDLE" << string(2, '\n');;
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	case SQL_ERROR:
		cout << "Error - SQL_ERROR" << string(2, '\n');
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	}

	if (SQLAllocHandle(SQL_HANDLE_STMT, SQLConnectionHandle, &SQLStatementHandle) == SQL_ERROR) {
		cout << left << setw(10) << "Error allocating SQL Statement Handle" << endl;
	}

	// Display the number of open exams 
	char SQLQuery[] = "Select count(*) From dbo.EWExam Where closed = '';";

	if (SQLExecDirect(SQLStatementHandle, (SQLCHAR*)SQLQuery, SQL_NTS) == SQL_ERROR) {
		cout << left << setw(10) << "Error executing SQL Statement Handle" << endl;
		showSQLErrorMsg(SQL_HANDLE_STMT, SQLStatementHandle);
	}
	else
	{
		while (SQLFetch(SQLStatementHandle) == SQL_SUCCESS) {
			SQLGetData(SQLStatementHandle, 1, SQL_C_DEFAULT, &count, sizeof(count) + 1, NULL);
		}
		cout << setw(15) << left << "Number of open exams: " << right << setw(4) << count << endl;
	}
	SQLFreeHandle(SQL_HANDLE_STMT, SQLStatementHandle);

	char clearOpt = 'N';
	cout << string(2, '\n');
	cout << "Enter Y to close all open exams, Z to exit: ";
	cin >> clearOpt;
	clearOpt = toupper(clearOpt);

	// Close all open exams
	if (clearOpt == 'Y') {
		cls();
		header();

		cout << endl << right << setw(20) << "Closing all open exams: ";

		char SQLQuery1[] = "Update dbo.EWExam Set closed='C' Where closed='';";

		if (SQLAllocHandle(SQL_HANDLE_STMT, SQLConnectionHandle, &SQLStatementHandle) == SQL_ERROR) {
			cout << left << setw(10) << "Error allocating SQL Statement Handle" << endl;
		}

		if (SQLExecDirect(SQLStatementHandle, (SQLCHAR*)SQLQuery1, SQL_NTS) == SQL_ERROR) {
			cout << left << setw(10) << "Error executing SQL Statement Handle" << endl;
			showSQLErrorMsg(SQL_HANDLE_STMT, SQLStatementHandle);
		}
		else
		{
			cout << right << setw(20) << "Successfully closed " << count << " open exams.";
		}
		pause();
	}

	if (clearOpt != 'Z' && clearOpt != 'Y') {
		cout << endl << " Invalid character entered. Try again." << endl;
		pause();
		closeAllOpenExams();
	}
	if (clearOpt == 'Z') {
		SQLmenu();
	}
}
