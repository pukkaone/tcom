// $Id$
#ifndef TCOMAPI_H
#define TCOMAPI_H

#pragma warning(disable: 4251)

#ifdef TCOM_EXPORTS
#define TCOM_API __declspec(dllexport)
#else
#define TCOM_API __declspec(dllimport)
#endif

#endif
