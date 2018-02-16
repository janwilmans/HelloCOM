// HelloCOMClient.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <iostream>
#include <system_error>

#import "../../HelloCOMObject/HelloCOMObject/Debug/HelloCOMObject.tlb" raw_interfaces_only no_smart_pointers named_guids


//class CAppModuleInitialization
//{
//public:
//    CAppModuleInitialization(CAppModule& module, HINSTANCE hInstance) :
//        m_module(module)
//    {
//        HRESULT hr = m_module.Init(nullptr, hInstance);
//        if (FAILED(hr))
//            Win32::ThrowWin32Error(hr, "CAppModule::Init");
//    }
//
//    ~CAppModuleInitialization()
//    {
//        m_module.Term();
//    }
//
//private:
//    CAppModule& m_module;
//};


std::wstring GetHresultMessage(HRESULT hr)
{
    _com_error err(hr);
    return err.ErrorMessage();
}

class HResultException : public std::system_error
{
public:
    HResultException(DWORD error, const std::string& what) :
        std::system_error(error, std::system_category(), what),
        m_what(GetHresultMessage(error))
    {
    }

    std::wstring GetMessage() const
    {
        return m_what;
    }
private:
    std::wstring m_what;
};

void CheckHr(HRESULT hr, const char* command, const char* file, int line)
{
    if (FAILED(hr)) throw HResultException(hr, command);
}

#define CHECKHR(x) CheckHr(x, #x, __FILE__, __LINE__)

void test()
{
    //CAppModuleInitialization moduleInit(_Module, hInstance);

    CHECKHR(OleInitialize(nullptr));

    //CLSID clsid;
    //IDispatch* pDispatch;
    //HRESULT nResult1 = CLSIDFromProgID(OLESTR("se.mysoft"), &clsid);
    //HRESULT nResult2 = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, __uuidof(HelloCOMObjectLib::IHello), (void **)&pDispatch);

    IUnknown* p_IUnknown;
    CHECKHR(CoCreateInstance(HelloCOMObjectLib::CLSID_HelloObject, NULL, CLSCTX_INPROC_SERVER, __uuidof(HelloCOMObjectLib::IHello), (void **)&p_IUnknown));

    //    MULTI_QI multi_qi[1];
    //auto hr = CoCreateInstanceEx(HelloCOMObjectLib::CLSID_HelloObject, NULL, CLSCTX_SERVER, NULL /*machinename*/, sizeof(multi_qi) / sizeof(multi_qi[0]), multi_qi);
    //s
}

int _tmain(int argc, _TCHAR* argv[])
{
    try
    {
        test();
    }
    catch (const HResultException& e)
    {
        std::cerr << "Exception: ";
        std::wcerr << e.GetMessage().c_str() << L"\n";
        std::cerr << " in '" << e.what() << "'\n\n";
    }
    catch (const std::exception& e)
    {
        std::cerr << "Exception: " << e.what() << "\n\n";
    }

    OleUninitialize();
	return 0;
}

//AccessMOT::AccessMOT(LPCWSTR nameMachine)
//{
//    // query machine object table on given machine
//    HRESULT result;
//    COSERVERINFO machine;
//    MULTI_QI qiMOT[1];
//
//    // connect to the machine object table (on given machine)
//    ZeroMemory(&machine, sizeof(machine));
//    ZeroMemory(qiMOT, sizeof(qiMOT));
//    if (nameMachine && wcslen(nameMachine)) 
//    {
//        machine.pwszName = const_cast<LPWSTR>(nameMachine);
//    }
//    qiMOT[0].pIID = &__uuidof(m_motInterface);
//    result = CoCreateInstanceEx(CLSID_MachineObjectTable, NULL, CLSCTX_SERVER, machine.pwszName ? &machine : NULL, sizeof(qiMOT) / sizeof(qiMOT[0]), qiMOT);
//    if (FAILED(result) || qiMOT[0].hr) 
//    {
//        // show information message in debug terminal
//        INFRAREPORT(Error, BindingQuery, L"AccesMOT", L"Failed to %s with '%s'\n", FAILED(result) ? L"connect to the MOT" : L"query Machine Object Table Interface", Infra::GetCOMErrorMessage(FAILED(result) ? result : qiMOT[0].hr).c_str());
//        // throw exception to show connection setup failed
//        throw MonikerIntermediateUnsupported(L"Failed to connect to the Machine Object Table. Error = '%s'", Infra::GetCOMErrorMessage(FAILED(result) ? result : qiMOT[0].hr).c_str());
//    }
//
//    // return requested interface
//    m_motInterface.Attach(reinterpret_cast<IMachineObjectTable*>(qiMOT[0].pItf));
//}
//
////CoCreateInstanceEx(CLSID_MachineObjectTable, NULL, CLSCTX_SERVER, machine.pwszName ? &machine : NULL, sizeof(qiMOT) / sizeof(qiMOT[0]), qiMOT);
////so you can even try to create object using simple CoCreate