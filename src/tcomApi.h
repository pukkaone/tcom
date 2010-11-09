// $Id: tcomApi.h,v 1.1 2000/04/22 21:39:36 chuang Exp $
#ifndef TCOMAPI_H
#define TCOMAPI_H

#pragma warning(disable: 4251)

#ifdef TCOM_EXPORTS
#define TCOM_API __declspec(dllexport)
#else
#define TCOM_API __declspec(dllimport)
#endif

#endif
