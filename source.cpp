#include <string>
#include <iostream>
#include <fstream>
#include <filesystem>
#include <vector>
#include <sstream>

#include <fcntl.h>
#include <io.h>


#include "excelmanager.h"

#include "format.h"

using namespace std;

void createOrEmptyDistDirectory(const string& distPath) {
    if (!filesystem::exists(distPath)) {
        filesystem::create_directory(distPath);
        wcout << "Dossier 'dist' créé avec succès !" << endl;
    }
    else {
        // Si le dossier existe, vide son contenu
        for (const auto& entry : filesystem::directory_iterator(distPath)) {
            filesystem::remove_all(entry);
        }
        wcout << "Contenu du dossier 'dist' vidé avec succès !" << endl;
    }
}



int main(int argc, char const* argv[])
{
    
    //_setmode(_fileno(stdout), _O_U16TEXT);
    //_setmode(_fileno(stdin), _O_U16TEXT);

    wcout << ITALIC << "Lancement du générateur de code :" << RESET << endl;
    const string distPath = "./dist";

    createOrEmptyDistDirectory(distPath);

    try {
        ExcelManager excelManager;
        vector<wstring> selectedFiles = excelManager.selectExcelFiles();

        if (!selectedFiles.empty()) {
            wcout << L"Fichiers sélectionnés : " << endl;
            for (const auto& file : selectedFiles) {
                wcout << file << endl;
            }
        }
        else {
            wcout << L"Sélection annulée par l'utilisateur." << endl;
        }
    }
    catch (const exception& e) {
        cerr << "Erreur : " << e.what() << endl;
    }

   
    std::cout << "Conversion du fichier Excel en CSV terminée." << std::endl;


    return EXIT_SUCCESS;
}