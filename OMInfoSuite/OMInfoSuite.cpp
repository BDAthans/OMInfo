
#include <iostream>
#include <iomanip>
#include <string>
#include <algorithm>
#include <fstream>
#include <atlstr.h>
#include <Shlobj.h>
#include <Windows.h>
#include <stdlib.h>
#include <dos.h>
#include <VersionHelpers.h>
#include <conio.h>

using namespace std;

#ifndef UNICODE  
typedef std::string String;
#else
typedef std::wstring String;
#endif

//System Variables
TCHAR wWinDir[80]; String WinDir;
TCHAR wDataDir[80]; String DataDir;
TCHAR wPgmsDir[80]; String PgmsDir;

//ADOConnection Variables
TCHAR wConnectThru[80]; String ConnectThru;
TCHAR wDatabaseName[80]; String DatabaseName;
TCHAR wDataSource[80]; String DataSource;

//INSTALL Variables
TCHAR wSQLbuild[80]; String SQLbuild;
TCHAR wInstalledVersion[80]; String InstalledVersion;
TCHAR wBuild[80]; String Build;
TCHAR wServicePack[80]; String ServicePack;

//Windows Specific Variables
String rUsername; //USERNAME
String rUserprofile; //USERPROFILE
String rLocalAppData; //LOCALAPPDATA
String rHostname; //COMPUTERNAME
String rSystemRoot; //SYSTEMROOT 

bool debugOn = false;
String runningVersion = "v0.0.15";

bool run = true;

int getOmate32();
int OMRunAsAdmin();
int duplicateINI();
void delTmpFiles();

void menu();
void header();
void cls();
void pause();
void adminPriv();
void winSvrChk();
int getSysInfo();
void resizeWindow();
void exit();
void logOutput();

/*
	FEATURES & NOTES TO ADD:

	- check for Officemate installed to more than one folder
		(Look for file in C:\Officemate AND C:\Omate32)
	- run all ~Reg .bat files for OM v12
	- export EnableLinkedConnections.reg and then run it for Windows 8+ only
	- clean up old Officemate installations and move them to C:\Old Officemate Installations
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
	//cout << string(26, '\n');
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
	cout << String(2, '\n');
	cout << setw(10) << left << "Option" << setw(15) << left << "Solutions";
	cout << endl << "--------------------------------------------------------------------------------";
	cout << endl << setw(10) << left << " A." << setw(40) << left << "Set OM\\EW Executables to run as Administrator (NOT WORKING YET)";
	cout << endl << setw(10) << left << " B." << setw(40) << left << "Delete .tmp files on C:\\ left from Reports";
	cout << endl << setw(10) << left << " C." << setw(40) << left << "Check for Duplicate Omate32.ini in common folders";
	cout << endl << setw(10) << left << " Z." << setw(40) << left << "Exit" << endl << endl;

	char menuopt;
	char confirmopt;

	cout << setw(10) << left << "Enter menu option from above to run solution: ";
	cin >> menuopt;
	menuopt = toupper(menuopt);

	//If menuopt is equal to Z it exits and doesn't do solution check
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
		delTmpFiles();
		break;
	case 'C':
		duplicateINI();
		break;
	case 'Z':
		exit();
		break;
	}
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

void exit() {
	cout << string(2, '\n') << "Press any Key to Exit... ";
	char a;
	a = _getch();
	exit(EXIT_SUCCESS);
}

void logOutput() {
	//Output info             1

	cout << setw(20) << left << "WinDir           = " << left << WinDir << endl;
	cout << setw(20) << left << "DataDir          = " << left << DataDir << endl;
	cout << setw(20) << left << "PgmsDir          = " << left << PgmsDir << endl << endl;
	cout << setw(20) << left << "ConnectThru      = " << left << ConnectThru << endl;
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
	/*
	- Set Omate.exe to run as Administrator for all users
	- Set ExamWriter.exe to run as Administrator for all users
	- Set Login.exe to run as Administrator for all users
	- Set Logindotnet.exe to run as Administrator for all users
	- Set HomeOffice.exe to run as Administrator for all users
	- Set Query.exe to run as Administrator for all users

	- set folder permissions for PgmsDir
	- set file permissions for Omate32.ini in WinDir

	- checks EnableLinkedConnections.reg value and verifies it is set correctly

	*/


	HKEY HKey;
	LPCSTR subk = TEXT("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers");
	LPCSTR kvalue = TEXT("C:\\Officemate\\Omate.exe");
	LPCSTR kdata = TEXT("~ RUNASADMIN");

	if (RegOpenKey(HKEY_LOCAL_MACHINE, subk, &HKey) != ERROR_SUCCESS)
	{
		cout << setw(10) << left << "Error opening Registry Key..." << endl;
	}
	else {

		if (RegSetValueEx(HKey, kvalue, 0, REG_SZ, (LPBYTE)kdata, strlen(kdata) * sizeof(char)) != ERROR_SUCCESS)
		{
			cout << setw(10) << left << "Error writing Registry value..." << endl;
		}
		else { cout << setw(10) << left << "Success writing Registry value..." << endl; }

		// More if RegSetValueEx to follow once it's working for the other executabes, other than Omate.exe

		LPCSTR kvalue = TEXT("C:\\Officemate\\Login.exe");
		if (RegSetValueEx(HKey, kvalue, 0, REG_SZ, (LPBYTE)kdata, strlen(kdata) * sizeof(char)) != ERROR_SUCCESS)
		{
			cout << setw(10) << left << "Error writing Registry value..." << endl;
		}
		else { cout << setw(10) << left << "Success writing Registry value..." << endl; }
	}

	RegCloseKey(HKey);
	pause();
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
	if (!vs.is_open()) {}
	else {
		cout << setw(10) << left << " WARNING: Duplicate Omate32.ini found in AppData\\Local\\VirtualStore\\Windows" << endl;
		fileCount++;
	}
	vs.close();

	tmp = "\\Windows\\Omate32.ini";
	string uppath = rUserprofile + tmp;
	up.open(uppath, ios::in);
	if (!up.is_open()) {}
	else {
		cout << setw(10) << left << " WARNING: Duplicate Omate32.ini found in Users Windows Folder" << endl;
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