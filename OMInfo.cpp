// OMInfo.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <iomanip>
#include <atlstr.h>
#include <string>
#include <iostream>
#include "Windows.h"

using namespace std;

//System
TCHAR wWinDir[80];
TCHAR wDataDir[80];
TCHAR wPgmsDir[80];

//ADOConnection
TCHAR wConnectThru[80];
TCHAR wDatabaseName[80];
TCHAR wDataSource[80];

//INSTALL
TCHAR wSQLbuild[80];
TCHAR wInstalledVersion[80];
TCHAR wServicePack[80];

int getOmate32();

int main()
{
	getOmate32();

	//Output info
	//_tprintf(TEXT("Key: %s\n"), wDataDir);
	wcout << "DataDir = " << wDataDir << endl;
	system("pause");

	     return 0;
}

int getOmate32() {

	//System Section
	GetPrivateProfileString(TEXT("System"), TEXT("WinDir"), TEXT("C:\\Windows"), wWinDir, 255, TEXT("Omate32.ini"));
	GetPrivateProfileString(TEXT("System"), TEXT("DataDir"), TEXT("C:\\Officemate\\Data"), wDataDir, 255, TEXT("Omate32.ini"));
	GetPrivateProfileString(TEXT("System"), TEXT("PgmsDir"), TEXT("C:\\Officemate"), wPgmsDir, 255, TEXT("Omate32.ini"));
	//ADOConnection Section
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("ConnectThru"), TEXT("MSDE"), wConnectThru, 255, TEXT("Omate32.ini"));
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("PgmsDir"), TEXT("C:\\Officemate\\Data"), wDatabaseName, 255, TEXT("Omate32.ini"));
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("PgmsDir"), TEXT("C:\\Officemate\\Data"), wDataSource, 255, TEXT("Omate32.ini"));
	//Install Section
	GetPrivateProfileString(TEXT("Install"), TEXT("PgmsDir"), TEXT("C:\\Officemate\\Data"), wSQLbuild, 255, TEXT("Omate32.ini"));
	GetPrivateProfileString(TEXT("Install"), TEXT("PgmsDir"), TEXT("C:\\Officemate\\Data"), wInstalledVersion, 255, TEXT("Omate32.ini"));
	GetPrivateProfileString(TEXT("Install"), TEXT("PgmsDir"), TEXT("C:\\Officemate\\Data"), wServicePack, 255, TEXT("Omate32.ini"));
	

	return 0;
}
