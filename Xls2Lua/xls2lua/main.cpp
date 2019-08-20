#include <iostream>
#include <sstream>
#include <fstream>
#include <vector>

#import "C:\Program Files\Common Files\System\ADO\msado15.dll" no_namespace rename("EOF", "adoEOF") rename("BOF", "adoBOF")

inline void TESTHR(HRESULT x) {if FAILED(x) _com_issue_error(x);};

std::wstring a2uc(const char * buffer, int len) {
	std::wstring newbuffer;
	int nChars = ::MultiByteToWideChar(CP_ACP, 0, buffer, len, NULL, 0);
	if (nChars == 0) return newbuffer;

	newbuffer.resize(nChars);
	::MultiByteToWideChar(CP_ACP, 0, buffer, len, const_cast<wchar_t*>(newbuffer.c_str()), nChars);
	return newbuffer;
}

std::string uc2u(const wchar_t* buffer, int len) {
	std::string newbuffer;
	int nChars = ::WideCharToMultiByte(CP_UTF8, 0, buffer, len, NULL, 0, NULL, NULL);
	if (nChars == 0) return newbuffer;

	newbuffer.resize(nChars);
	::WideCharToMultiByte(CP_UTF8, 0, buffer, len, const_cast<char*>(newbuffer.c_str()), nChars, NULL, NULL);
	return newbuffer;
}

std::wstring a2uc(const std::string& str) {
	return a2uc(str.c_str(), (int)str.size());
}

std::string uc2u(const std::wstring& str) {
	return uc2u(str.c_str(), (int)str.size());
}

std::string a2u(const std::string& str) {
	return uc2u(a2uc(str));
}

std::string makeConnStr(std::string filename, bool header = true) {
    std::stringstream stream;
    std::string hdr = header ? "YES" : "NO";    
	if (!filename.empty()) {
		if (*--filename.end() == 'x') {
			stream << "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" << filename << ";Extended Properties=\"Excel 12.0 Xml;IMEX=1;HDR=" << hdr << "\"";
		} else {
			stream << "Provider='Microsoft.JET.OLEDB.4.0';Data Source=" << filename << ";Extended Properties=\"Excel 8.0;IMEX=1;HDR=" << hdr << "\"";
		}
	}
    return stream.str();
}

std::string sqlSelectSheet(_bstr_t connStr, int sheetIndex) {
    _ConnectionPtr pCon = NULL;
    _RecordsetPtr pSchema = NULL;
    TESTHR(pCon.CreateInstance(__uuidof(Connection)));
    TESTHR(pCon->Open(connStr, "", "", NULL));    
    pSchema = pCon->OpenSchema(adSchemaTables); 
    for(int i = 0; i < sheetIndex; ++i) pSchema->MoveNext();
    std::string sheetName = (char*)(_bstr_t)pSchema->Fields->GetItem("TABLE_NAME")->Value.bstrVal;
    std::stringstream stream;
    stream << "SELECT * FROM [" << sheetName << "]";
    return stream.str();
}


void export_lua(std::string excel, std::string luaname, std::string fname, int sheetIndex) {
	try {
		_RecordsetPtr pRec = NULL;
		_RecordsetPtr pSchema = NULL;
		_bstr_t connStr(makeConnStr(excel, false).c_str());
		TESTHR(pRec.CreateInstance(__uuidof(Recordset)));
		TESTHR(pRec->Open(sqlSelectSheet(connStr, sheetIndex).c_str(), connStr, adOpenStatic, adLockReadOnly, adCmdText));

		size_t line = 0;
		std::stringstream output;
		std::vector<std::string> header, typer;
		output << "--[[" << luaname << "\n";
		output << "--]]" << "\n\n";
		output << "local " << fname << "= \n{\n";

		while (!pRec->adoEOF) {
			size_t col = pRec->Fields->GetCount();
			if (line < 3) {
				if (line == 0) {
					for (long i = 0; i < col; ++i) {
						_variant_t v = pRec->Fields->GetItem(i)->Value;
						header.push_back(v.vt == VT_BSTR ? (char*)(_bstr_t)v.bstrVal : "");
					}
				} else if(line == 1){
					for (long i = 0; i < col; ++i) {
						_variant_t v = pRec->Fields->GetItem(i)->Value;
						typer.push_back(v.vt == VT_BSTR ? (char*)(_bstr_t)v.bstrVal : "");
					}
				}
				line++;
				pRec->MoveNext();
				continue;
			}
			for (long j = 0; j < col; ++j)
			{
				_variant_t value = pRec->Fields->GetItem(j)->Value;
				std::stringstream temp;
				if (j == 0) {
					if (header[j] == "id") {
						output << "\t" << "[" << (char*)(_bstr_t)value.bstrVal << "] = { ";
						continue;
					} else {
						output << "\t{ ";
					}
				}
				if (header[j] == "" || value.vt == VT_EMPTY || value.vt == VT_NULL) {
					//δ���ñ�ͷ���ֶ��Լ�����Ϊ�յ��ֶβ�����
					continue;
				}
				output << header[j] << " = ";
				if (value.vt == VT_R8) {
					output << value.dblVal << ", ";
				}
				else if (value.vt == VT_BSTR) {
					if (typer[j] != "string") {
						output << (char*)(_bstr_t)value.bstrVal << ", ";
					} else {
						output << "'" << (char*)(_bstr_t)value.bstrVal << "', ";
					}
				}
			}
			output << "},\n";
			pRec->MoveNext();
		}
		output << "}\n\n";
		output << "return " << fname;

		std::fstream outputFile;
		outputFile.open(luaname, std::ios::out | std::ios::trunc);
		outputFile << a2u(output.str());
		outputFile.flush();
		outputFile.close();
    } catch(_com_error &e) {        
        _bstr_t bstrDescription(e.Description());      
        CharToOem(bstrDescription, bstrDescription);
        std::cout << bstrDescription << std::endl;
    }  
}

int main(int argc, char **argv) {
	printf("Excel file(xls/xlsx) convert lua tool, Version 0.1.0\n\n");
	printf("********************************************************************\n");
	printf("������ʹ��ע�����\n");
	printf("1���뽫��Ҫת��lua���õ�xls�ļ��ŵ���������Ŀ¼�µ�xlsĿ¼��\n");
	printf("2��ÿ��xls�ļ�ֻ֧��ת��һ��sheet����������ΪSheet1��\n");
	printf("3��Excel���ñ��һ��Ϊ�ֶ������ڶ���Ϊ�ֶ����ͣ�stringΪ�ַ���������Ĭ�����֣���������Ϊ�ֶν������ӵ����п�ʼΪ���ݡ�\n");
	printf("4���ֶν����������������һ���ֶ���Ϊ�գ�����ֶβ�������lua��\n");
	printf("5��������lua�ļ��ڱ�Ŀ¼�µ�lua�ļ�������û�и��ļ������Լ������ļ��С�\n");
	printf("6����ͷ�ֶ��������Դ�д��ĸ��ͷ������֣���F1��F2������Excel�Ĺؼ��֡�\n");
	printf("7�����ݸ��Ӷ�����ֵ���жϸñ�ĵ�������˵�һ�в������ÿ�ֵ��\n");
	printf("8�������ʾȱ������������������װOffice����汾��\n");
	printf("9�����ϣ������{[key] = {...}}��ʽlua�����һ�ֶ�����Ϊid��\n");
	printf("********************************************************************\n\n\n");

	printf("��ʼ�����ļ������Եȣ�\n");

	if (argc < 2) {
		printf("�������ԣ�������Դ�ļ�������\n");
		return 0;
	}

	HANDLE hFile;
	WIN32_FIND_DATA pNextInfo;            
	std::stringstream search;
	search << argv[1] << "\\*.xls*";
	hFile = FindFirstFile(search.str().c_str(), &pNextInfo);
	if (hFile == INVALID_HANDLE_VALUE) {
		return 0;
	}
	if (FAILED(::CoInitialize(NULL))) {
		return 0;
	}
	do {
		std::string filename = pNextInfo.cFileName;
		size_t pos = filename.find(".");
		filename = filename.substr(0, pos);
		std::stringstream excelname, luaname;
		excelname << argv[1] << "\\" << pNextInfo.cFileName;
		luaname << (argv[2] ? argv[2] : argv[1]) << "\\" << filename << ".lua";
		export_lua(excelname.str(), luaname.str(), filename, 0);
		printf("%s�����ɹ���\n", pNextInfo.cFileName);
	} while (FindNextFile(hFile, &pNextInfo));

	::CoUninitialize();
	getchar();

	return 1;
}
