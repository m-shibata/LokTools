/**
 * lok_keyevent - input keyevent and paste label
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
#include <mutex>
#include <condition_variable>
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

    int inputKey(char col, char row, std::string label) {
        if (!_doc || label.empty()) return -1;

        if (_doc->getDocumentType() != LOK_DOCTYPE_SPREADSHEET) {
            std::cerr << "Error: Is not Calc file (type = ";
            std::cerr << _doc->getDocumentType() << ")" << std::endl;
            return -1;
        }

        /* GoTo the celll */
        for (int i = 0; i < col - 'A'; ++i) {
            _doc->postKeyEvent(LOK_KEYEVENT_KEYINPUT, 0, 1027);
            _doc->postKeyEvent(LOK_KEYEVENT_KEYUP, 0, 1027);
        }
        for (int i = 0; i < row - '1'; ++i) {
            _doc->postKeyEvent(LOK_KEYEVENT_KEYINPUT, 0, 1024);
            _doc->postKeyEvent(LOK_KEYEVENT_KEYUP, 0, 1024);
        }

        /* Paste label */
        if (!_doc->paste("text/plain;charset=utf-8",
                         label.c_str(), label.size())) {
            std::cerr << "Error: Failed to paste: " << label << std::endl;
        }

        /* GoTo A1 */
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
    std::cerr << " [-p path-of-libreoffice] [-r row] [-c column]";
    std::cerr << " CalcFile Label" << std::endl;
}

int main(int argc, char **argv)
{
    int opt;
    std::string lo_path = LokTools::LO_PATH;
    char column = 'A';
    char row = '1';
    while ((opt = getopt(argc, argv, "p:c:r:")) != -1) {
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
        default:
            usage(argv[0]);
            exit(EXIT_FAILURE);
        }
    }

    if (argc - optind < 2) {
        usage(argv[0]);
        exit(EXIT_FAILURE);
    }
    const char *calc_file = argv[optind++];
    const char *label = argv[optind];


    try {
        LokTools lok(lo_path);

        if (lok.open(calc_file)) {
            std::cerr << "Error: Failed to open" << std::endl;
            exit(EXIT_FAILURE);
        }

        if (lok.inputKey(column, row, label)) {
            std::cerr << "Error: Failed to input keys" << std::endl;
            exit(EXIT_FAILURE);
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
