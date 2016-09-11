/**
 * lok_convert - converter between LibreOffice document formats
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
#include <LibreOfficeKit/LibreOfficeKit.hxx>

namespace lok_tools {
    const std::string LO_PATH = "/usr/lib/libreoffice/program";
}

void usage(const char *name)
{
    std::cerr << "Usage: " << name;
    std::cerr << " [-p path-of-libreoffice] FROM TO" << std::endl;
}

int main(int argc, char **argv)
{
    int opt;
    std::string lo_path = lok_tools::LO_PATH;
    while ((opt = getopt(argc, argv, "p:")) != -1) {
        switch (opt) {
        case 'p':
            lo_path = std::string(optarg);
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
    const char *from_file = argv[optind++];
    const char *to_file = argv[optind];


    try {
        lok::Office *lo = lok::lok_cpp_init(lo_path.c_str());
        if (!lo) {
            std::cerr << "Error: Failed to initialise LibreOfficeKit";
            std::cerr << std::endl;
            exit(EXIT_FAILURE);
        }

        lok::Document *doc = lo->documentLoad(from_file);
        if (!doc) {
            std::cerr << "Error: Failed to load document: ";
            std::cerr << lo->getError() << std::endl;
            delete lo;
            exit(EXIT_FAILURE);
        }

        if (!doc->saveAs(to_file)) {
            std::cerr << "Error: Failed to save document: ";
            std::cerr << lo->getError() << std::endl;
            delete doc;
            delete lo;
            exit(EXIT_FAILURE);
        }

        delete doc;
        delete lo;
    } catch (const std::exception & e) {
        std::cerr << "Error: " << e.what() << std::endl;
        exit(EXIT_FAILURE);
    }

    return 0;
}
