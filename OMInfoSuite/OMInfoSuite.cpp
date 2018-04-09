
#include <iostream>
#include <iomanip>
#include <string>
#include <algorithm>
#include <atlstr.h>
#include <Shlobj.h>
#include <Windows.h>

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
TCHAR wServicePack[80];
String ServicePack;

bool showLog = true;
string runningVersion = "v1.0.0";

int getOmate32();
int setOMEWRunAsAdmin();
void header();
void exit();
void logOutput();

int main()
{
	header();
	getOmate32();
	if (showLog == true) { logOutput(); }
	setOMEWRunAsAdmin();
	

	exit();
	return 0;
}

void header() {
	cout << "Eyefinity Officemate Suite Information and Diagnostics Tool " << runningVersion << endl;
	cout << "--------------------------------------------------------------------------------" << string(1, '\n');
}

void exit() {
	cout << string(2, '\n');
	system("pause");
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
	cout << setw(20) << left << "ServicePack      = " << left << ServicePack << endl;

	//SQLbuild comparison is working
	cout << endl << "Expected SQLbuild: '" << SQLbuild << "'" << endl;
	if (SQLbuild == "12.0.2") {
		cout << "SQLbuild is matching " << SQLbuild << endl;
	}

	//PgmsDir comparison is not working
	cout << endl << "Expected PgmsDir: '" << PgmsDir << "'" << endl;
	if (PgmsDir == "C:\\OFFICEMATE") {
		cout << "PgmsDir is matching " << PgmsDir << endl;
	}
}	

int getOmate32() {
	cout << "Gathering Information from Omate32.ini..." << string(2, '\n');

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
	SQLbuild = wSQLbuild;
	GetPrivateProfileString(TEXT("Install"), TEXT("InstalledVersion"), TEXT(""), wInstalledVersion, 255, TEXT("Omate32.ini"));
	InstalledVersion = wInstalledVersion;
	GetPrivateProfileString(TEXT("Install"), TEXT("Service Pack"), TEXT(""), wServicePack, 255, TEXT("Omate32.ini"));
	ServicePack = wServicePack;

	//Converts Omate32.ini strings to UPPER case for comparison
	transform(WinDir.begin(), WinDir.end(), WinDir.begin(), (int(*)(int))toupper);
	transform(DataDir.begin(), DataDir.end(), DataDir.begin(), (int(*)(int))toupper);
	transform(PgmsDir.begin(), PgmsDir.end(), PgmsDir.begin(), (int(*)(int))toupper);

	transform(ConnectThru.begin(), ConnectThru.end(), ConnectThru.begin(), (int(*)(int))toupper);
	transform(DatabaseName.begin(), DatabaseName.end(), DatabaseName.begin(), (int(*)(int))toupper);
	transform(DataSource.begin(), DataSource.end(), DataSource.begin(), (int(*)(int))toupper);

	transform(SQLbuild.begin(), SQLbuild.end(), SQLbuild.begin(), (int(*)(int))toupper);
	transform(InstalledVersion.begin(), InstalledVersion.end(), InstalledVersion.begin(), (int(*)(int))toupper);
	transform(ServicePack.begin(), ServicePack.end(), ServicePack.begin(), (int(*)(int))toupper);

	return 0;
}

int setOMEWRunAsAdmin() {

	return 0;
}
