/* >>> Pictographics Intl. Corp. Confidential and Proprietary <<<
* This work contains valuable confidential and proprietary information.
* Disclosure, use or reproduction without the written authorization of
* Pictographics is prohibited. This unpublished work by Pictographics
* is protected by the laws of the United States and other countries. If
* publication of the work should occur the following notice shall apply:
* "Copyright (c) 1990 - 1999 Pictographics. All Rights Reserved"
*/
/*
* Name: Picto.h
*
* Author: B. J. Lindbloom
*
* Purpose:
* Constants and typedefs common to Pictographics software.
*/
#ifndef PICTO_H_
#define PICTO_H_
/* - - - - - - - - - - Constants - - - - - - - - - - */
#define CE_OK 0 /* ok status */
#define CE_ARGERROR -1 /* bad arguments to a function */
#define CE_GENERROR -2 /* general error (nothing else fits) */
#define CE_MALLOCERROR -3 /* memory allocation failure */
#define CE_OPENERROR -4 /* file open error */
#define CE_READERROR -5 /* file read error */
#define CE_WRITEERROR -6 /* file write error */
#define CE_SEEKERROR -7 /* file seek error */
#define CE_FORMATERROR -8 /* data format error */
#define CE_UNSUPERROR -9 /* unsupported feature or format */
#define CE_ABORT -10 /* user initiated abort */
#define CE_TIMEOUTERROR -11 /* time out of some process */
#ifndef TRUE
#define TRUE 1 /* true */
#undef FALSE
#define FALSE 0 /* false */
#endif
#define RL_UNKNOWN 0 /* unknown color data type or data format */
/* Image file formats */
#define RL_TIFF 1 /* TIFF format (8-bits per channel) */
#define RL_TARGA 2 /* TARGA format */
#define RL_SCITEXCT 3 /* Scitex CT format */
#define RL_EPS 4 /* EPS format */
#define RL_DCS 5 /* DCS format */
#define RL_PCD 6 /* Kodak Photo CD format */
#define RL_BMP 7 /* Windows Bitmap format */
#define RL_CINEON 8 /* Kodak Cineon format */
#define RL_JPEG 9 /* JPEG format */
#define RL_TIFF16 10 /* TIFF format (16-bits per channel) */
/* Color types */
#define RL_RGB 1 /* RGB data */
#define RL_CMYK 4 /* CMYK data */
#define RL_GRAY 5 /* grayscale data */
#define RL_LAB 7 /* CIELAB color data */
#define RL_YCC8 8 /* PhotoYcc 8-bit color data */
#define MAX_NCHAN 4 /* maximum permitted number of channels
per scanline */
#define RL_DEFAULTRES 300 /* default resolution (pixels per inch) */
#define CC_INPUT 0 /* referring to ColorCircuit inputs */
#define CC_OUTPUT 1 /* referring to ColorCircuit outputs */
/* Math constants */
#ifndef MATH_CONSTANTS
#ifndef M_PI
#define M_PI 3.141592653589793 /* pi */
#endif
#ifndef M_LN2
#define M_LN2 0.693147180559945 /* ln(2) */
#endif
#ifndef M_E
#define M_E 2.718281828459045 /* e */
#endif
#ifndef M_GR
#define M_GR 1.618033988749895 /* golden ratio */
#endif
#ifndef M_R2
#define M_R2 1.414213562373095 /* sqrt(2) */
#endif
#ifndef M_R3
#define M_R3 1.732050807568877 /* sqrt(3) */
#endif
#ifndef M_LOG2
#define M_LOG2 0.301029995663981 /* log10(2) */
#endif
#endif /* #ifndef MATH_CONSTANTS */
/* - - - - - - - - - - Macros - - - - - - - - - - */
#define COLOR3COPY(a,b) (b)[0] = (a)[0]; \
(b)[1] = (a)[1]; \
(b)[2] = (a)[2]
#define COLOR3MAKE(a,b,c,d) (a)[0] = (b); \
(a)[1] = (c); \
(a)[2] = (d)
#define COLOR3SET(a,b) (a)[0] = (a)[1] = (a)[2] = (b)
#define COLOR4COPY(a,b) (b)[0] = (a)[0]; \
(b)[1] = (a)[1]; \
(b)[2] = (a)[2]; \
(b)[3] = (a)[3]
#define COLOR4MAKE(a,b,c,d,e) (a)[0] = (b); \
(a)[1] = (c); \
( a)[2] = (d); \
(a)[3] = (e)
#define COLOR4SET(a,b) (a)[0] = (a)[1] = (a)[2] = (a)[3] = (b)
#define CLAMP(v,min,max) if ((v) < (min)) \
(v) = (min); \
else if ((v) > (max)) \
(v) = (max)
#define LERP(a,b,c) (((b) - (a)) * (c) + (a))
#define ILERP(a,b,c) ((b) - (a)) / ((c) - (a))
#define SUM3(a) ((a)[0] + (a)[1] + (a)[2])
#define SUM4(a) ((a)[0] + (a)[1] + (a)[2] + (a)[3])
#define MAX2(a,b) ((a) >= (b) ? (a) : (b))
#define MIN2(a,b) ((a) <= (b) ? (a) : (b))
#define MAX3(a,b,c) (MAX2((a), MAX2((b), (c))))
#define MIN3(a,b,c) (MIN2((a), MIN2((b), (c))))
#define IS_ZERO(v,eps) (((v) < (eps)) && ((v) > -(eps)))
#define NOT_ZERO(v,eps) (((v) > (eps)) || ((v) < -(eps)))
/* - - - - - - - - - - Typedefs - - - - - - - - - - */
#ifndef CLStat
#if defined(WIN32) && defined(DLL)
#define CLStat __declspec(dllimport) int
#else
#define CLStat int
#endif
#endif
#ifndef UCHAR_
typedef unsigned char uchar;
#define UCHAR_
#endif
#ifndef USHORT_
typedef unsigned short ushort;
#define USHORT_
#endif
#ifndef UINT_
typedef unsigned int uint;
#define UINT_
#endif
#ifndef ULONG_
typedef unsigned long ulong;
#define ULONG_
#endif
#ifndef BOOLEAN_
typedef unsigned char boolean;
#define BOOLEAN_
#endif
#ifndef FLT_
typedef float Flt;
#define FLT_
#endif
#ifndef FLT2_
typedef double Flt2;
#define FLT2_
#endif
#ifndef FLTX_
typedef long double FltX;
#define FLTX_
#endif
#ifndef RFLT_
#ifdef CL_DOUBLE_IS_LONG
typedef long double RFlt;
#else
typedef double RFlt;
#endif
#define RFLT_
#endif
#ifndef MFLT_
#ifdef CL_DOUBLE_IS_LONG
typedef long double MFlt;
#else
typedef double MFlt;
#endif
#define MFLT_
#endif
#ifndef RATIONAL_
typedef ulong rational[2];
#define RATIONAL_
#endif
#ifndef ColorCircuit_
/* Note: Library users should consider ColorCircuit structures to be read-only! */
typedef struct ColorCircuit {
int nIn; /* number of input channels */
int nOut; /* number of output channels */
int iType; /* color type of input */
int oType; /* color type of output */
int ccBPC; /* ColorCircuit internal number of bytes per
channel {1, 2} */
int usrIBPC; /* user number of bytes per input channel {1, 2} */
int usrOBPC; /* user number of bytes per output channel {1, 2} */
void *data; /* pointer to other data */
} ColorCircuit;
#define ColorCircuit_
#endif
#endif /* #ifndef PICTO_H_ */