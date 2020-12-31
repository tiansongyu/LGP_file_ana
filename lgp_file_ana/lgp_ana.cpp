#include "lgp_ana.h"

lgp_ana::lgp_ana()
{

}
LGP_DATA lgp_ana::find_data(const char* _file_dir,const char* _file_name)
{
    qDebug() << QString("%1").arg(_file_dir);

    LGP_DATA tmp_lgp;
    const char* file_name = new char[strlen(_file_name)+1];
    memcpy((void*)file_name,_file_name,strlen(_file_name)+1);

    const char* file_dir = new char[strlen(_file_dir)+1];
    memcpy((void*)file_dir,_file_dir,strlen(_file_dir)+1);

    int number = 0;
    int current = 0;
    char signature[8];
    // = { 0x00,0x00,0xC8,0x42,0xcd,0xcc,0x4c,0x3e };
    signature[0] = 0x00;
    signature[1] = 0x00;
    signature[2] = 0xC8;
    signature[3] = 0x42;
    signature[4] = 0xcd;
    signature[5] = 0xcc;
    signature[6] = 0x4c;
    signature[7] = 0x3e;

    std::ifstream myfile(file_dir, std::ios::binary);
    myfile.seekg(0, std::ios_base::end);
    const int myfileLen = myfile.tellg();
    myfile.seekg(0, myfile.beg);

    if (!myfile.is_open())
    {
        qDebug()<< "打开文件失败" ;
    }
    if(myfileLen < 1024 * 1024 )
    {
        tmp_lgp.file_time = QString::fromUtf8(file_name);
        tmp_lgp.ptop = QString("没有数据");
        tmp_lgp.change = QString("没有数据");
        tmp_lgp.frequency = QString("没有数据");
        tmp_lgp.acceleration = QString("没有数据");

        return tmp_lgp;
    }
    char* lgd_file = new char[myfileLen];
    memset(lgd_file, 0, myfileLen);
    qDebug() << "读取了 " << myfileLen << " 个字节... ";

    myfile.read(lgd_file, myfileLen);
    if (myfile)
        qDebug() << "所有字节扫描完毕." ;
    else
        qDebug() <<"error: only " << myfile.gcount() << " could be read";
    char* orgin = lgd_file;
    while (number < 2)
    {
        if (memcmp(orgin, signature, 8) == 0)
        {
            number++;
        }
        else
        {
            current++;
            if (current > myfileLen - 0x2c)
            {
                std::cout << "没有找到" << file_name << "目标特征" << std::endl;
                return tmp_lgp;
            }
        }
        orgin++;
    }
    current++;
    //输出加速度/g	叶片微应变	叶片振幅（p-p）/mm	叶片频率/Hz

    tmp_lgp.file_time = file_name;
    tmp_lgp.acceleration = QString("%1").arg(*(float*)(lgd_file + current + acceleration_offset));
    tmp_lgp.frequency = QString("%1").arg(*(float*)(lgd_file + current + frequency_offset));
    qDebug() << "成功找到文件" << file_name << "的目标特征";
    //qDebug() << "该文件的加速度的值为  " << set_float(4) * (float*)(lgd_file + current + acceleration_offset) ;
    //qDebug() << "该文件的频率的值为    " << set_float(4) * (float*)(lgd_file + current + frequency_offset) ;
    //计算应力
    float* stress_value = new float[256];
    memcpy((void*)stress_value, lgd_file + stress, 1024);
    float min = stress_value[0];
    float max = stress_value[0];

    for (int i = 0; i < 256; i++)
    {
        if (stress_value[i] < min)
            min = stress_value[i];
        if (stress_value[i] > max)
            max = stress_value[i];
    }
    float result = max - min;

    qDebug() <<  "该文件的应力的值为    " << result ;
    qDebug() <<  "该文件的微应变的值为    " << result / 200 * 1000 * 10000 ;
    tmp_lgp.change = QString("%1").arg(result / 200 * 1000 * 10000);
    float* amplitude_value = new float[256];
    memcpy((void*)amplitude_value, lgd_file + amplitude, 1024);
    float min_amplitude = amplitude_value[0];
    float max_amplitude = amplitude_value[0];
    for (int i = 0; i < 256; i++)
    {
        if (amplitude_value[i] < min_amplitude)
            min_amplitude = amplitude_value[i];
        if (amplitude_value[i] > max_amplitude)
            max_amplitude = amplitude_value[i];
    }
    float result_amplitude_value = max_amplitude - min_amplitude;

    tmp_lgp.ptop =  QString("%1").arg(result_amplitude_value);


    delete[] lgd_file;
    delete[] file_dir;
    delete[] file_name;
    return tmp_lgp;
}
