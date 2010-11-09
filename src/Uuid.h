// $Id$
#ifndef UUID_H
#define UUID_H

#include <string.h>
#include <string>
#include <comdef.h>
#include "tcomApi.h"

// This class wraps a UUID to provide convenience functions.

class TCOM_API Uuid
{
    UUID m_uuid;

public:
    // Construct from UUID.
    Uuid (const UUID &uuid):
        m_uuid(uuid)
    { }

    // less than operator
    bool operator< (const Uuid &rhs) const
    { return memcmp(&m_uuid, &rhs.m_uuid, sizeof(UUID)) < 0; }

    // equals operator
    bool operator== (const Uuid &rhs) const
    { return memcmp(&m_uuid, &rhs.m_uuid, sizeof(UUID)) == 0; }

    // Return string representation.
    std::string toString() const;
};

#endif
