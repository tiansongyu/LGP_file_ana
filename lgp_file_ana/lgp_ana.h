#ifndef LGP_ANA_H
#define LGP_ANA_H

#include <qstring.h>
#include <qdebug.h>

#include <Windows.h>

#include <fstream>
#include <iostream>
#include <string.h>
#include <iomanip>
#include <vector>
#include <string>
#include <ShlObj.h>

#define _CRT_SECURE_NO_WARNINGS

#define _CRT_SECURE_NO_DEPRECATE
#define _SCL_SECURE_NO_DEPRECATE

#define set_float(x) std::setiosflags(std::ios::fixed) << std::setprecision(x) <<

#define acceleration_offset 0xf0 + 4
#define frequency_offset + 0xf0 + 4 - 0x30

#define stress 0x158888
#define amplitude 0xcf678


struct LGP_DATA
{
    QString file_time;
    QString acceleration;
    QString change;
    QString ptop;
    QString frequency;
};


class lgp_ana
{
public:
    lgp_ana();
    LGP_DATA find_data(const char* file_dir);

};

#endif // LGP_ANA_H
