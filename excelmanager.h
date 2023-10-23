#pragma once
#include <iostream>
#include <vector>
#include <string>

#include <windows.h>
#include <commdlg.h> 
#include <shobjidl.h>
#include <objbase.h> 
#include <algorithm>
#include <fstream>

using namespace std;

class ExcelManager {
public:
    ExcelManager();
    ~ExcelManager();

    vector<wstring> selectExcelFiles();
private:
    bool checkValidExtension(const std::wstring& filePath);
};