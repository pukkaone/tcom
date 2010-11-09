// $Id: Uuid.cpp,v 1.2 2000/04/20 18:37:40 chuang Exp $
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
