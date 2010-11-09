// $Id$
#ifndef REGISTRYKEY_H
#define REGISTRYKEY_H

#include <stdexcept>
#include <string>
#define WIN32_LEAN_AND_MEAN
#include <windows.h>

// This class represents a registry key.

class RegistryKey
{
    HKEY m_hkey;

    // Open registry key.
    void open(HKEY hkey, const std::string &subkeyName);

public:
    RegistryKey(HKEY hkey, const std::string &subkeyName);
    RegistryKey(const RegistryKey &key, const std::string &subkeyName);
    ~RegistryKey();

    // Get name of subkey under this key.
    std::string subkeyName(int index);

    // Get data for default value under this key.
    std::string value();

    // Get data for value under this key.
    std::string value(const char *valueName);
};

#endif
