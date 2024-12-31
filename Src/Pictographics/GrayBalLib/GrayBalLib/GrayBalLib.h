Chapter 5: GrayBalLib.h Listing
/* >>> Pictographics Intl. Corp. Confidential and Proprietary <<<
* This work contains valuable confidential and proprietary information.
* Disclosure, use or reproduction without the written authorization of
* Pictographics Intl. Corp. is prohibited. This unpublished work by
* Pictographics Intl. Corp. is protected by the laws of the United States
* and other countries. If publication of the work should occur the
* following notice shall apply:
* "Copyright (c) 1990 - 2003 Pictographics Intl. Corp. All Rights Reserved"
*/
/*
* Name: GrayBalLib.h
*
* Author: B. J. Lindbloom
* Tracy Finks
*
* Purpose:
* Header file for GrayBalLib routines.
*
* Notes: All densities are expressed in integer form, scaled by 1000.
* For example:
*
* densityInt = (int)(densityFloat * 1000.0 + 0.5);
*
* Gray Balance test patterns must have step sizes of equal integer increments,
* so 2, 4, 6, 16, 18, 52, 86, or 256 patch step wedges are supported.
*/
#ifndef GRAYBALLIB_H_
#define GRAYBALLIB_H_
/* - - - - - includes - - - - - */
#include "Picto.h"
/* - - - - - defines - - - - - */
/* possible values for 'modMethod' in 'Target' data structure */
#define GB_TARGET_MOD_METHOD_LINEAR 0 /* linearly transform targets to
actual media range */
#define GB_TARGET_MOD_METHOD_CLAMP 1 /* match actual targets, if
possible, else clamp */
#define GB_TARGET_MOD_METHOD_LINEAR_MINMAX 2 /* linearly transform targets to
media min/max range */
/* - - - - - typedefs - - - - - */
typedef int GbColorInt[3]; /* RGB for density, pixel value or exposure set */
typedef struct GbTarget { /* specification of a set of target densities */
int nTargetSamples;/* number of samples in set */
GbColorInt *digital; /* set of 'nTargetSamples' digital values [0, 255] */
GbColorInt *density; /* set of 'nTargetSamples' densities */
int modMethod; /* method to be used to modify target densities */
} GbTarget;
typedef struct GbDensReport { /* report summarizing density analysis */
int dMaxError; /* error in maximum density */
int dMinError; /* error in minimum density */
int grayBalRMS; /* RMS of gray balance error */
int grayBalMaxErr; /* maximum gray balance error */
int grayBalStep; /* step of gray scale that had maximum error [0, 9] */
int badMeasMaxErr; /* largest "irregular appearing" measurement, CMY */
int badMeasStep; /* which step had the above error [0, nMeas - 1], CMY */
} GbDensReport;
typedef struct GbMap {
Flt lut[3][256]; /* contents in range [0.0, 1.0] */
} GbMap;
/* - - - - - function prototypes - - - - - */
#ifdef __cplusplus
extern "C" {
#endif
CLStat
GbExamineDensities(
int nMeas, /* number of measurements in array measDens */
/* (one of 8, 13, 16, 24, 64, 72, 208, 344, 1024) */
GbColorInt *measDens,/* pointer to array of nMeas density measurements */
GbTarget *target, /* pointer to target density specification */
GbDensReport *report /* pointer to where report should go */
);
CLStat
GbRefineExposureSet( /* refine current exposure set */
GbColorInt inES, /* current exposure set */
GbColorInt outES, /* new exposure set returned here (may be same as inES)
*/
GbColorInt stepSize, /* step sizes (must be > 0) */
GbColorInt targetDens, /* desired target densities */
int nMeas, /* number of measurements in measDens (13 or 30) */
GbColorInt *measDens /* pointer to array of nMeas density measurements */
);
CLStat
GbRefineGrayMap( /* refine current gray balance map */
GbMap *inGMap, /* pointer to input gray balance map */
GbMap *outGMap, /* pointer to output gray balance map */
/* (may be same as inGMap) */
int nMeas, /* number of measurements in measDens */
/* (4 * steps in test pattern) */
/* (must be one of 8, 16, 24, 64, 72, 208, 344, 1024) */
GbColorInt *measDens,/* pointer to array of nMeas density measurements*/
GbTarget *target /* pointer to target density specification */
);
#ifdef __cplusplus
}
#endif
#endif /* #ifndef GRAYBALLIB_H_ */