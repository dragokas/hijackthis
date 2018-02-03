This is a "ready-to-use" VS 2015 solution of the PCRE2 library (v10.23), designed for use with legacy VB6 projects.  The library should work on Win XP through Win 10 (but please note that XP support is not well-tested).

The solution builds the PCRE-16 library, specifically, which allows it to be used against VB6 BSTRs with minimal work.  PCRE's built-in Unicode support and JIT features are enabled.  

Release DLLs are built using stdcall instead of cdecl, and function names *remain mangled* (e.g. "pcre2_compile_context_create" resolves to "_pcre2_compile_context_create@4").

For a ready-to-use VB6 wrapper that makes use of this DLL, please visit https://github.com/jpbro/VbPcre2

PCRE is BSD-licensed.  The full source code and API documentation are available at http://www.pcre.org/

For the source code of pcre2-16.dll look in repository by tannerhelland: https://github.com/tannerhelland/PCRE2-VB6-DLL/releases

--------

HiJackThis note:

This program is an essential part of HiJackThis Fork project resources.

---------

Checksum:

pcre2-16.dll
SHA1: 76E72667E79615669079BCE9CF5B627B74C532F6
SHA256: c304ba9a4d54f3a4cb130c7849fdfc482691fc7a526c9f73d7c9dbb3a0899307

Digitally signed by Alex Dragokas (using self-signed certificate).
Certificate's thumbprint should be: 05F1F2D5BA84CDD6866B37AB342969515E3D912E

