This is a "ready-to-use" VS 2015 solution of the PCRE2 library (v10.23), designed for use with legacy VB6 projects.  The library should work on Win XP through Win 10 (but please note that XP support is not well-tested).

The solution builds the PCRE-16 library, specifically, which allows it to be used against VB6 BSTRs with minimal work.  PCRE's built-in Unicode support and JIT features are enabled.  

Release DLLs are built using stdcall instead of cdecl, and function names *remain mangled* (e.g. "pcre2_compile_context_create" resolves to "_pcre2_compile_context_create@4").

For a ready-to-use VB6 wrapper that makes use of this DLL, please visit https://github.com/jpbro/VbPcre2

PCRE is BSD-licensed.  The full source code and API documentation are available at http://www.pcre.org/
