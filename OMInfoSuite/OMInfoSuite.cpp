
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

using namespace std;

#ifndef UNICODE  
typedef std::string String;
#else
typedef std::wstring String;
#endif

//System Variables
TCHAR wWinDir[80];
String WinDir;
TCHAR wDataDir[80];
String DataDir;
TCHAR wPgmsDir[80];
String PgmsDir;

//ADOConnection Variables
TCHAR wConnectThru[80];
String ConnectThru;
TCHAR wDatabaseName[80];
String DatabaseName;
TCHAR wDataSource[80];
String DataSource;

//INSTALL Variables
TCHAR wSQLbuild[80];
String SQLbuild;
TCHAR wInstalledVersion[80];
String InstalledVersion;
TCHAR wBuild[80];
String Build;
TCHAR wServicePack[80];
String ServicePack;

String rUsername;
String rUserprofile;

bool showLog = true;
string runningVersion = "v1.0.0";

bool run = true;

int getOmate32();
int OMPermissions();
int duplicateINI();

void menu();
void header();
void cls();
void adminPriv();
void getSysInfo();
void exit();
void logOutput();

int main()
{
	adminPriv();
	while (run == true) {
		cls();
		header();
		getOmate32();
		if (showLog == true) { logOutput(); }
		menu();
	}
	exit();
	return 0;
}

void header() {
	cout << "Eyefinity Officemate Suite Information and Diagnostics Tool " << runningVersion << endl;
	cout << "--------------------------------------------------------------------------------" << string(1, '\n');
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

void adminPriv() {
	if (IsUserAnAdmin() == false) {
		header();
		cout << String(2, '\n') << "NEED TO RUN AS ADMINISTRATOR. CONTACT IT";
		exit();
	}
}

void menu() {
	cout << String(1, '\n');
	cout << endl << setw(10) << left << "Option" << setw(15) << left << "Solutions";
	cout << endl << "--------------------------------------------------------------------------------";
	cout << endl << setw(10) << left << "A." << setw(40) << left << "Set OM\EW Executables to run as Administrator";
	cout << endl << setw(10) << left << "Z." << setw(40) << left << "Exit" << endl << endl;

	char menuopt;
	char confirmopt;

	cout << setw(10) << left << "Enter menu option from above to run solution: ";
	cin >> menuopt;
	menuopt = toupper(menuopt);

	//If menuopt is equal to Z it exits and doesn't do solution check
	if (menuopt != 'Z') {
		cout << setw(10) << left << "Are you sure you want to run solution - " << menuopt << "? Y or N:";
		cin >> confirmopt;
		confirmopt = toupper(confirmopt);
		if (confirmopt != 'Y') { 
			main();;
		}
		//Need to add check so that if confirmopt is not Y or N, it restarts 
		if ((confirmopt != 'N' && confirmopt == 'Y') || (confirmopt == 'N' && confirmopt != 'Y'))
		{
			/*Further check here*/
		}
		else { main(); }
	}
	
	switch(menuopt)
	{
	case 'A' : 
		OMPermissions();
		break;
	case 'Z' : 
		exit();
		break;
	}
}

void getSysInfo() {
	/*
	rUsername = getenv("username");
	rUserprofile = getenv("userprofile");
	
	char ptr[80];
	char *path;
	int i = 0;
	unsigned int len;
	 Get the current path environment 
	getenv_s(&len, ptr, 80, "PATH");  */

}

void exit() {
	cout << string(2, '\n');
	cout << setw(10) << "Closing. Press Any Key to Continue";
	Sleep(2000);
	exit(EXIT_SUCCESS);
}

void logOutput() {
	//Output info

	/*
	wcout << setw(20) << left << "WinDir           = " << left << wWinDir << endl;
	wcout << setw(20) << left << "DataDir          = " << left << wDataDir << endl;
	wcout << setw(20) << left << "PgmsDir          = " << left << wPgmsDir << endl << endl;
	wcout << setw(20) << left << "ConnectThru      = " << left << wConnectThru << endl;
	wcout << setw(20) << left << "DatabaseName     = " << left << wDatabaseName << endl;
	wcout << setw(20) << left << "DataSource       = " << left << wDataSource << endl;
	wcout << setw(20) << left << "SQL_Build        = " << left << wSQLbuild << endl << endl;
	wcout << setw(20) << left << "InstalledVersion = " << left << wInstalledVersion << endl;
	wcout << setw(20) << left << "ServicePack      = " << left << wServicePack << endl;
	*/
	
	cout << setw(20) << left << "WinDir           = " << left << WinDir << endl;
	cout << setw(20) << left << "DataDir          = " << left << DataDir << endl;
	cout << setw(20) << left << "PgmsDir          = " << left << PgmsDir << endl << endl;
	cout << setw(20) << left << "ConnectThru      = " << left << ConnectThru << endl;
	cout << setw(20) << left << "DatabaseName     = " << left << DatabaseName << endl;
	cout << setw(20) << left << "DataSource       = " << left << DataSource << endl;
	cout << setw(20) << left << "SQL_Build        = " << left << SQLbuild << endl << endl;
	cout << setw(20) << left << "InstalledVersion = " << left << InstalledVersion << endl;
	cout << setw(20) << left << "Build = " << left << Build << endl;
	cout << setw(20) << left << "ServicePack      = " << left << ServicePack << endl;

	//SQLbuild comparison is working
	cout << endl << "Expected SQLbuild: '" << SQLbuild << "'" << endl;
	if (SQLbuild == "12.0.2") {
		cout << "SQLbuild is matching " << SQLbuild << endl;
	}

	//PgmsDir comparison is not working
	cout << endl << "Expected PgmsDir: '" << PgmsDir << "'" << endl;
	if (PgmsDir == "C:\\OFFICEMATEV14") {
		cout << "PgmsDir is matching " << PgmsDir << endl;
	}
}	

int getOmate32() {
	cout << "Gathering Information from Omate32.ini..." << string(2, '\n');

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

	if (DataDir == "INI NOT FOUND") {
		cout << setw(10) << left << "Omate32.ini not found in C:\Windows" << endl;
		cout << setw(10) << left << "Please correct Omate32.ini to proceed" << endl;
		Sleep(2000); 
		exit();
	}

	return 0;
}

int OMPermissions() {
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
	return 0;
}

int duplicateINI() {
	/*
	- check for Officemate installed to more than one location???
	- check for files in C:\Officemate or C:\Omate32, or else where

	- check if duplicate Omate32.ini exists in %appdata%/../local/virtualstore
	- check if duplicate Omate32.ini exists in %userprofile%/Windows

	*/
	return 0;
}