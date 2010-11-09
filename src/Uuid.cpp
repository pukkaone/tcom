// $Id: Uuid.cpp 5 2005-02-16 14:57:24Z cthuang $
#include "Uuid.h"

std::string
Uuid::toString () const
{
    unsigned char *str;
    if (UuidToString(const_cast<UUID *>(&m_uuid), &str) != RPC_S_OK) {
        return std::string();
    }
    std::string result(reinterpret_cast<char *>(str));
    RpcStringFree(&str);
    return result;
}
