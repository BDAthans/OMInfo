#include "stdafx.h"
#include <iomanip>
#include <atlstr.h>
#include <string>
#include <iostream>
#include "Windows.h"
#include "Shlobj.h"

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
String SQLBuild;
TCHAR wInstalledVersion[80];
String InstalledVersion;
TCHAR wServicePack[80];
String ServicePack;

boolean showLog = true;
string runningVersion = "v1.0.0";

int getOmate32();
void header();
void output();

int main()
{
	header();
	getOmate32();
	if (showLog == true) {output();}
	
	cout << string(1, '\n');
	system("pause");
	     return 0;
}

void header() {
	cout << "Eyefinity Officemate Suite Information and Diagnostics Tool " << runningVersion << endl;
	if (IsUserAnAdmin() == false) { cout << "Run As Administrator" << endl;  }
	cout << "--------------------------------------------------------------------------------" << string(1, '\n');
}

int getOmate32() {
	cout << endl << "Gathering Information from Omate32.ini..." << string(2, '\n');

	//System Section
	GetPrivateProfileString(TEXT("System"), TEXT("WinDir"), TEXT(""), wWinDir, 255, TEXT("Omate32.ini"));
	WinDir = wWinDir;
	GetPrivateProfileString(TEXT("System"), TEXT("DataDir"), TEXT(""), wDataDir, 255, TEXT("Omate32.ini"));
	DataDir = wDataDir;
	GetPrivateProfileString(TEXT("System"), TEXT("PgmsDir"), TEXT(""), wPgmsDir, 255, TEXT("Omate32.ini"));
	PgmsDir = wPgmsDir;

	//ADOConnection Section
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("ConnectThru"), TEXT(""), wConnectThru, 255, TEXT("Omate32.ini"));
	ConnectThru = wConnectThru;
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("DatabaseName"), TEXT(""), wDatabaseName, 255, TEXT("Omate32.ini"));
	DatabaseName = wDatabaseName;
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("DataSource"), TEXT(""), wDataSource, 255, TEXT("Omate32.ini"));
	DataSource = wDataSource;

	//Install Section
	GetPrivateProfileString(TEXT("Install"), TEXT("SQL_Build"), TEXT(""), wSQLbuild, 255, TEXT("Omate32.ini"));
	SQLBuild = wSQLbuild;
	GetPrivateProfileString(TEXT("Install"), TEXT("InstalledVersion"), TEXT(""), wInstalledVersion, 255, TEXT("Omate32.ini"));
	InstalledVersion = wInstalledVersion;
	GetPrivateProfileString(TEXT("Install"), TEXT("Service Pack"), TEXT(""), wServicePack, 255, TEXT("Omate32.ini"));
	ServicePack = wServicePack;

	return 0;
}

void output() {
	//Output info
	//_tprintf(TEXT("Key: %s\n"), wDataDir);

	wcout << setw(20) << left << "WinDir           = "  << left << WinDir << endl;
	wcout << setw(20) << left << "DataDir          = " << left << DataDir << endl;
	wcout << setw(20) << left << "PgmsDir          = " << left << PgmsDir << endl <<endl;
	wcout << setw(20) << left << "ConnectThru      = " << left << ConnectThru << endl;
	wcout << setw(20) << left << "DatabaseName     = " << left << DatabaseName << endl;
	wcout << setw(20) << left << "DataSource       = " << left << DataSource << endl;
	wcout << setw(20) << left << "SQL_Build        = " << left << SQLBuild << endl << endl;
	wcout << setw(20) << left << "InstalledVersion = " << left << InstalledVersion << endl;
	wcout << setw(20) << left << "ServicePack      = " << left << ServicePack << endl;
}
