/* config.h for CMake builds */

/* #undef HAVE_DIRENT_H */
#define HAVE_INTTYPES_H 1    
#define HAVE_STDINT_H 1                                                   
#define HAVE_STRERROR 1
#define HAVE_SYS_STAT_H 1
#define HAVE_SYS_TYPES_H 1
/* #undef HAVE_UNISTD_H */
#define HAVE_WINDOWS_H 1

/* #undef HAVE_BCOPY */
#define HAVE_MEMMOVE 1

/* #undef PCRE2_STATIC */

/* #undef SUPPORT_PCRE2_8 */
#define SUPPORT_PCRE2_16 1
/* #undef SUPPORT_PCRE2_32 */
/* #undef PCRE2_DEBUG */

/* #undef SUPPORT_LIBBZ2 */
/* #undef SUPPORT_LIBEDIT */
/* #undef SUPPORT_LIBREADLINE */
/* #undef SUPPORT_LIBZ */

#define SUPPORT_JIT 1
/* #undef SUPPORT_PCRE2GREP_JIT */
#define SUPPORT_UNICODE 1
/* #undef SUPPORT_VALGRIND */

/* #undef BSR_ANYCRLF */
/* #undef EBCDIC */
/* #undef EBCDIC_NL25 */
#define HEAP_MATCH_RECURSE 1
/* #undef NEVER_BACKSLASH_C */

#define LINK_SIZE		2
#define MATCH_LIMIT		10000000
#define MATCH_LIMIT_RECURSION	MATCH_LIMIT
#define NEWLINE_DEFAULT         5
#define PARENS_NEST_LIMIT       250
#define PCRE2GREP_BUFSIZE       20480
#define PCRE2GREP_MAX_BUFSIZE   1048576

#define MAX_NAME_SIZE	32
#define MAX_NAME_COUNT	10000

/* end config.h for CMake builds */
