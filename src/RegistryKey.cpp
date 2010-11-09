// $Id$
#include "RegistryKey.h"

void
RegistryKey::open (HKEY hkey, const std::string &subkeyName)
{
    LONG result = RegOpenKeyEx(
        hkey,
        subkeyName.c_str(),
        0,
        KEY_READ,
        &m_hkey);
    if (result != ERROR_SUCCESS) {
        throw std::runtime_error("cannot read registry key " + subkeyName);
    }
}

RegistryKey::RegistryKey (HKEY hkey, const std::string &subkeyName)
{
    open(hkey, subkeyName);
}

RegistryKey::RegistryKey (const RegistryKey &key,
                          const std::string &subkeyName)
{
    open(key.m_hkey, subkeyName);
}

RegistryKey::~RegistryKey ()
{
    RegCloseKey(m_hkey);
}

std::string
RegistryKey::subkeyName (int index)
{
    char name[256];
    DWORD size = sizeof(name);
    FILETIME lastWriteTime;

    LONG result = RegEnumKeyEx(
        m_hkey,
        index,
        name,
        &size,
        NULL,
        NULL,
        NULL,
        &lastWriteTime);
    if (result != ERROR_SUCCESS) {
        throw std::runtime_error("RegEnumKeyEx");
    }

    return std::string(name);
}

std::string
RegistryKey::value ()
{
    return value("");
}

std::string
RegistryKey::value (const char *valueName)
{
    BYTE data[256];
    DWORD size = sizeof(data);

    LONG result = RegQueryValueEx(
        m_hkey,
        valueName,
        NULL,
        NULL,
        data,
        &size);
    if (result != ERROR_SUCCESS) {
        throw std::runtime_error("RegQueryValueEx");
    }

    return std::string(reinterpret_cast<char *>(data));
}
