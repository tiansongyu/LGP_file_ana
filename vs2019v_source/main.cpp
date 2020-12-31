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

TCHAR szPathName[MAX_PATH];

char* get_dir()
{
	BROWSEINFO bInfo = { 0 };
	bInfo.hwndOwner = GetForegroundWindow();//父窗口    
	bInfo.lpszTitle = TEXT("浏览文件夹");
	bInfo.ulFlags = BIF_RETURNONLYFSDIRS | BIF_USENEWUI/*包含一个编辑框 用户可以手动填写路径 对话框可以调整大小之类的..*/ |
		BIF_UAHINT/*带TIPS提示*/ | BIF_NONEWFOLDERBUTTON /*不带新建文件夹按钮*/;
	LPITEMIDLIST lpDlist ;
	lpDlist = SHBrowseForFolder(&bInfo);
	if (lpDlist != NULL)
		{
			SHGetPathFromIDList(lpDlist, szPathName);\
			for (int i = 0; szPathName[i] != NULL; i++)
			{
				if (szPathName[i] == '\\')
					szPathName[i] = '/';
			}
			//MessageBox(NULL, szPathName, "Dir Name", MB_OK);
			//std::cout << szPathName << std::endl;
		}
		else
		{
			printf("user cancle\n");
		}
	return szPathName;
}

void find_data(const char* file_dir)
{
	const char* file_name = file_dir;
	int number = 0;
	int current = 0;
	char signature[] = { 0x00,0x00,0xc8,0x42,0xcd,0xcc,0x4c,0x3e };
	std::ifstream myfile(file_name, std::ios::binary);
	myfile.seekg(0, std::ios_base::end);
	const int myfileLen = myfile.tellg();
	myfile.seekg(0, myfile.beg);

	if (!myfile.is_open())
	{
		std::cout << "打开文件失败" << std::endl;
	}
	char* lgd_file = new char[myfileLen];
	memset(lgd_file, 0, myfileLen);
	std::cout << "读取了 " << myfileLen << " 个字节... ";

	myfile.read(lgd_file, myfileLen);
	if (myfile)
		std::cout << "所有字节扫描完毕." << std::endl;
	else
		std::cout << "error: only " << myfile.gcount() << " could be read";
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
				return ;
			}
		}
		orgin++;
	}
	current++;
	std::cout << "成功找到文件" << file_name << "的目标特征" << std::endl;
	std::cout << "该文件的加速度的值为  " << set_float(4) * (float*)(lgd_file + current + acceleration_offset) << std::endl;
	std::cout << "该文件的频率的值为    " << set_float(4) * (float*)(lgd_file + current + frequency_offset) << std::endl;
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

	std::cout << "该文件的应力的值为    " << result << std::endl;
	std::cout << "该文件的微应变的值为    " << result / 200 * 1000 * 10000 << std::endl;

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

	std::cout << "该文件的振幅的值为    " << set_float(7) result_amplitude_value << std::endl;

	std::cout << std::endl;
	delete lgd_file;
}
int main()
{
	std::cout << "选择.lgp文件所在目录" << std::endl;
	std::vector<std::string> vfile_name;
	char* dir = new char[50];
	dir = get_dir();

	int file_number = 0;
	WIN32_FIND_DATA p;
	char* d = new char[40];
	sprintf(d, "%s/*.lgp", dir);
	HANDLE h = FindFirstFile(d, &p);
	if (h == INVALID_HANDLE_VALUE)
	{
		std::cout << "没有找到任何 .lgp文件,请选择有.lgp文件存在的目录" << std::endl;
		system("pause");
		return 0;
	}
	std::cout << "从" << dir << "/中找到如下文件" << std::endl;
	puts(p.cFileName);
	vfile_name.push_back(std::string(p.cFileName));

	file_number++;
	while (FindNextFile(h, &p))
	{
		puts(p.cFileName);
		vfile_name.push_back(std::string(p.cFileName));
		file_number++;
	}
	std::cout << "共找到" << file_number << "个文件" << std::endl << std::endl;
	int tmp_number = file_number;

	std::ofstream output("output.txt");
	
	while (file_number > 0)
	{
		char* tmp_dir = new char[100];
		sprintf(tmp_dir, "%s/%s", dir, vfile_name[tmp_number - file_number].c_str());
		std::cout << tmp_dir << std::endl;
		find_data(tmp_dir);
		file_number--;
	}
	system("pause");
	return 0;
}