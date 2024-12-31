/*
 *	Header file for GrayBalLib test programs.
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "GrayBalLib.h"

#define	BUF_SIZE		128
#define	MY_TARGET_MOD_METHOD	GB_TARGET_MOD_METHOD_LINEAR_MINMAX

void	GetNonCommentLine(FILE *fd, char *buf);
void	PrintReport(GbDensReport *report);
int	ReadMapFile(char *mapFileName, GbMap *map);
int	ReadMeasDensFile(char *inMeasFile, int nMeasMax, GbColorInt *measDens, int *nMeas);
int	ReadTargetFile(char *targetFileName, GbTarget *target);
int	WriteMapFile(char *mapFileName, GbMap *map);
int	WriteTargetFile(char *targetFileName, GbTarget *target);