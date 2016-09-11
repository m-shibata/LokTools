/**
 * lok_screenshot - take and paste screenshots
 *
 * Written in 2016 by Mitsuya Shibata <mty.shibata@gmail.com>
 *
 * To the extent possible under law, the author(s) have dedicated all copyright
 * and related and neighboring rights to this software to the public domain
 * worldwide. This software is distributed without any warranty.
 *
 * You should have received a copy of the CC0 Public Domain Dedication along
 * with this software. If not, see
 * <http://creativecommons.org/publicdomain/zero/1.0/>.
 */
#include <stdlib.h>
#include <unistd.h>
#include <iostream>
#include <fstream>
#include <mutex>
#include <condition_variable>
#include <thread>
#define LOK_USE_UNSTABLE_API
#include <LibreOfficeKit/LibreOfficeKit.hxx>
#include <LibreOfficeKit/LibreOfficeKitEnums.h>

class LokTools {
public:
    static const std::string LO_PATH;

    LokTools(std::string path) :
        isUnoCompleted(false),
        _lo(NULL),
        _doc(NULL) {
        _lo = lok::lok_cpp_init(path.c_str());
        if (!_lo)
            throw;
    };

    ~LokTools() {
        if (_doc) delete _doc;
        if (_lo) delete _lo;
    };

    int open(std::string file) {
        _doc = _lo->documentLoad(file.c_str());
        if (!_doc) {
            std::cerr << "Error: Failed to load document: ";
            std::cerr << _lo->getError() << std::endl;
            return -1;
        }

        _doc->registerCallback(docCallback, this);
        return 0;
    };

    static void docCallback(int type, const char* payload, void* data) {
        LokTools* self = reinterpret_cast<LokTools*>(data);

        switch (type) {
        case LOK_CALLBACK_UNO_COMMAND_RESULT:
            {
                std::unique_lock<std::mutex> lock(self->unoMtx);
                self->isUnoCompleted = true;
            }
            self->unoCv.notify_one();
            break;
        default:
            ; /* do nothing */
        }
    };

    void postUnoCommand(const char *cmd, const char *args = NULL,
                        bool notify = false) {
        isUnoCompleted = false;
        _doc->postUnoCommand(cmd, args, notify);

        if (notify) {
            std::unique_lock<std::mutex> lock(unoMtx);
            unoCv.wait(lock, [this]{ return isUnoCompleted;});
        }
    };

    int takeScreenshot(char col, char row, int sheet) {
        if (!_doc) return -1;

        if (_doc->getDocumentType() != LOK_DOCTYPE_SPREADSHEET) {
            std::cerr << "Error: Is not Calc file (type = ";
            std::cerr << _doc->getDocumentType() << ")" << std::endl;
            return -1;
        }

        if (sheet >= _doc->getParts()) {
            std::string sheetName = "Sheet" + std::to_string(sheet+1);
            std::string sheetCmd = ".uno:Add";
            std::string sheetArg = "{"
                "\"Name\":{"
                    "\"type\":\"string\","
                    "\"value\":\"" + sheetName + "\""
                "}}";
            postUnoCommand(sheetCmd.c_str(), sheetArg.c_str(), true);
        }
        _doc->setPart(sheet);

        gotoCell(std::string(1, col), std::string(1, row));

        /* Take screenshot */
        std::string fileName = "/tmp/lok_tools_screen.png";
        std::string cmdName = "import -window root " + fileName;
        if (system(cmdName.c_str())) {
            std::cerr << "Error: Failed to take screenshot" << std::endl;
        }

        insertGraphic(fileName);
        gotoCell("A", "1");

        return 0;
    };

    void gotoCell(const std::string &col, const std::string &row) {
        std::string command = ".uno:GoToCell";
        std::string arguments = "{"
            "\"ToPoint\":{"
                "\"type\":\"string\","
                "\"value\":\"$" + col + "$" + row + "\""
            "}}";
        postUnoCommand(command.c_str(), arguments.c_str(), true);
    };

    void insertGraphic(const std::string &file) {
        std::string scheme = "file://";
        std::string command = ".uno:InsertGraphic";
        std::string argument = "{"
            "\"FileName\":{"
                "\"type\":\"string\","
                "\"value\":\"" + scheme + file + "\""
            "}}";
        postUnoCommand(command.c_str(), argument.c_str(), true);
    };

    int save() {
        postUnoCommand(".uno:Save", NULL, true);
        return 0;
    };

    std::mutex unoMtx;
    std::condition_variable unoCv;
    bool isUnoCompleted;

private:
    lok::Office *_lo;
    lok::Document *_doc;
};
const std::string LokTools::LO_PATH = "/usr/lib/libreoffice/program";

void usage(const char *name)
{
    std::cerr << "Usage: " << name;
    std::cerr << " [-p path-of-libreoffice] [-c column] [-r row]";
    std::cerr << " [-s sheets] [-i interval] CalcFile" << std::endl;
}

int main(int argc, char **argv)
{
    int opt;
    std::string lo_path = LokTools::LO_PATH;
    char column = 'A';
    char row = '1';
    int sheets = 1;
    int interval = 5;
    while ((opt = getopt(argc, argv, "p:c:r:s:i:")) != -1) {
        switch (opt) {
        case 'p':
            lo_path = std::string(optarg);
            break;
        case 'c':
            if (isalpha(*optarg))
                column = toupper(*optarg);
            break;
        case 'r':
            if (isdigit(*optarg))
                row = *optarg;
            break;
        case 's':
            sheets = atoi(optarg);
            break;
        case 'i':
            interval = atoi(optarg);
            break;
        default:
            usage(argv[0]);
            exit(EXIT_FAILURE);
        }
    }

    if (argc - optind < 1) {
        usage(argv[0]);
        exit(EXIT_FAILURE);
    }
    const char *calc_file = argv[optind++];


    try {
        LokTools lok(lo_path);

        if (lok.open(calc_file)) {
            std::cerr << "Error: Failed to open" << std::endl;
            exit(EXIT_FAILURE);
        }

        for (int i = 0; i < sheets; ++i) {
            std::cout << "Take screenshot for Sheet" << i+1 << std::endl;

            if (lok.takeScreenshot(column, row, i)) {
                std::cerr << "Error: Failed to take screenshot" << std::endl;
                exit(EXIT_FAILURE);
            }

            std::this_thread::sleep_for(std::chrono::seconds(interval));
        }

        if (lok.save()) {
            std::cerr << "Error: Failed to save document" << std::endl;
            exit(EXIT_FAILURE);
        }

    } catch (const std::exception & e) {
        std::cerr << "Error: " << e.what() << std::endl;
        exit(EXIT_FAILURE);
    }

    return 0;
}
