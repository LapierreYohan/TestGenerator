#include "excelmanager.h"


ExcelManager::ExcelManager() {
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(hr)) {
        throw runtime_error("Erreur lors de l'initialisation de COM.");
    }
}

ExcelManager::~ExcelManager() {
    CoUninitialize();
}

vector<wstring> ExcelManager::selectExcelFiles() {

    std::vector<std::wstring> selectedFiles;
    bool validSelection = false;

    do {
        IFileOpenDialog* pFileOpen;
        HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL, IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));
        if (FAILED(hr)) {
            throw std::runtime_error("Erreur lors de la création de la boîte de dialogue.");
        }

        COMDLG_FILTERSPEC filterSpec[] = { { L"Excel Files", L"*.xlsx;*.xls" }, { L"CSV Files", L"*.csv" }, { L"All Files", L"*.*" } };
        pFileOpen->SetFileTypes(3, filterSpec);
        pFileOpen->SetTitle(L"Sélectionnez un ou plusieurs fichiers Excel ou CSV");
        pFileOpen->SetOptions(FOS_ALLOWMULTISELECT);

        // Message indiquant qu'on ne souhaite que des fichiers Excel ou CSV
        pFileOpen->SetOkButtonLabel(L"Sélectionner");
        pFileOpen->SetFileNameLabel(L"Veuillez sélectionner un fichier Excel ou CSV");

        hr = pFileOpen->Show(NULL);
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED)) {
            // L'utilisateur a annulé la sélection
            break;
        }
        else if (SUCCEEDED(hr)) {
            IShellItemArray* pResults;
            pFileOpen->GetResults(&pResults);

            DWORD fileCount;
            pResults->GetCount(&fileCount);

            for (DWORD i = 0; i < fileCount; ++i) {
                PWSTR pszFilePath;
                IShellItem* pItem;
                pResults->GetItemAt(i, &pItem);
                pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);

                // Vérifier l'extension du fichier (Excel ou CSV)
                std::wstring filePath(pszFilePath);
                if (checkValidExtension(filePath)) {
                    selectedFiles.emplace_back(filePath);
                }
                else {
                    // Afficher un message d'erreur sous forme de fenêtre pop-up
                    MessageBox(NULL, L"Impossible de choisir ce type de fichier.\nVeuillez sélectionner un fichier Excel ou CSV valide.", L"Erreur de sélection", MB_OK | MB_ICONERROR);
                    validSelection = false; // Réinitialiser le drapeau de sélection valide
                    break; // Sortir de la boucle pour redemander à l'utilisateur de choisir un fichier
                }

                CoTaskMemFree(pszFilePath);
                pItem->Release();
            }

            pResults->Release();

            if (!selectedFiles.empty()) {
                validSelection = true; // Indiquer que la sélection est valide
            }
        }

        pFileOpen->Release();
    } while (!validSelection); // Continuer tant que l'utilisateur n'a pas choisi de fichier valide ou a annulé

    return selectedFiles;
}

bool ExcelManager::checkValidExtension(const wstring& filePath) {
    // Vérifier l'extension du fichier
    wstring extension = filePath.substr(filePath.find_last_of(L".") + 1);
    transform(extension.begin(), extension.end(), extension.begin(), ::towlower);

    return (extension == L"xlsx" || extension == L"xls" || extension == L"csv");
}
