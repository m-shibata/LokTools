/**
 * lok_split - extractor for Impress pages
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
#include <vector>
#include <fstream>
#include <png++/png.hpp>
#define LOK_USE_UNSTABLE_API
#include <LibreOfficeKit/LibreOfficeKit.hxx>
#include <LibreOfficeKit/LibreOfficeKitEnums.h>

namespace lok_tools {
    const std::string LO_PATH = "/usr/lib/libreoffice/program";
}

void usage(const char *name)
{
    std::cerr << "Usage: " << name;
    std::cerr << " [-p path-of-libreoffice] [-d dpi] ImpressFile BaseName";
    std::cerr << std::endl;
}

int splitImpress(lok::Document *doc, std::string base, int dpi)
{
    if (!doc || base.empty() || dpi == 0) return -1;

    if (doc->getDocumentType() != LOK_DOCTYPE_PRESENTATION) {
        std::cerr << "Error: Is not Impress file (type = ";
        std::cerr << doc->getDocumentType() << ")" << std::endl;
        return -1;
    }

    std::string extname = ".png";
    for (int i = 0; i < doc->getParts(); ++i) {
        long pageWidth, pageHeight;

        doc->setPart(i);
        doc->getDocumentSize(&pageWidth, &pageHeight);

        /* 1 Twips = 1/1440 inch */
        int canvasWidth = pageWidth * dpi / 1440;
        int canvasHeight = pageHeight * dpi / 1440;
        std::vector<unsigned char> pixmap(canvasWidth * canvasHeight * 4);
        doc->paintTile(pixmap.data(), canvasWidth, canvasHeight, 0, 0,
                       pageWidth, pageHeight);

        png::image <png::rgba_pixel> image(canvasWidth, canvasHeight);
        for (int x = 0; x < canvasWidth; ++x) {
            for (int y = 0; y < canvasHeight; ++y) {
                image[y][x].red     = pixmap[(canvasWidth * y + x) * 4];
                image[y][x].green   = pixmap[(canvasWidth * y + x) * 4 + 1];
                image[y][x].blue    = pixmap[(canvasWidth * y + x) * 4 + 2];
                image[y][x].alpha   = pixmap[(canvasWidth * y + x) * 4 + 3];
            }
        }
        std::string filename = base + std::to_string(i) + extname;
        image.write(filename);
    }

    return 0;
}

int main(int argc, char **argv)
{
    int opt;
    std::string lo_path = lok_tools::LO_PATH;
    int dpi = 96;
    while ((opt = getopt(argc, argv, "p:d:")) != -1) {
        switch (opt) {
        case 'p':
            lo_path = std::string(optarg);
            break;
        case 'd':
            dpi = atoi(optarg);
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
    const char *base_name = argv[optind];


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

        if (splitImpress(doc, base_name, dpi)) {
            std::cerr << "Error: Failed to split document" << std::endl;
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
