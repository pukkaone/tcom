# $Id$

include_external_msproject(tcom
        ${CMAKE_CURRENT_SOURCE_DIR}/tcom.vcproj)
include_external_msproject(dllserver
        ${CMAKE_CURRENT_SOURCE_DIR}/dllserver.vcproj
        tcom)
include_external_msproject(exeserver
        ${CMAKE_CURRENT_SOURCE_DIR}/exeserver.vcproj
        tcom)
include_external_msproject(TclScript
        ${CMAKE_CURRENT_SOURCE_DIR}/TclScript.vcproj
        tcom)
