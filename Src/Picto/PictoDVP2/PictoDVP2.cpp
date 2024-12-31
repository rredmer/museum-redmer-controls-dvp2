/****************************************************************************
**                                                                         **
**  System......: DVP2 Pictographics Support Application                   **
**                                                                         **
**  Module......: PictoDVP2.cpp                                            **
**                                                                         **
**  Description.: This application provides a command-line interface to the**
**                Pictographics Gray Balance Library for photographic paper**
**                linearization.                                           **
**                                                                         **
**  History.....:                                                          **
**                03/20/04 RDR Designed & Programmed first release.        **
**                                                                         **
** (c) 2004 Redmer Controls Inc.  All Rights Reserved.                     **
****************************************************************************/
#include "stdafx.h"				// Standard Microsoft C++ header
#include "tGrayBalLib.h"		// Pictographics Gray Balance Library
#define	GB_MAX_TP_STEPS		256	// maximum number of steps allowed in gray balance test pattern

static void	ErrorCleanup(GbTarget *target, GbMap *map, GbColorInt *measDens);
static void	ExitNow(int exitStatus);


void main(int argc, char **argv)
{
	char buf[256];
	char path[200];
	GbColorInt *measDens;
	GbDensReport report;
	int status, nMeas;
	GbTarget target;
	GbMap *map;


	printf("\nDVP2 Pictographics GrayBalLib Interface Version 1.00, March 20, 2004");
	printf("\n--------------------------------------------------------------------");

	// Initialize.
 	map = NULL;
	target.digital = NULL;
	target.density = NULL;
	measDens = NULL;

	
	// Get the path to setup files from the command-line
	if (argc < 2) {
		printf("\nPlease pass path to input files as command-line argument.\n\n");
		exit(-1);
	}
	sprintf(path, "%s", argv[1]);
	printf("\nSetting path to: %s", path);


	// Read the input map file.
	sprintf(buf, "%s%s", path, "PictoIN.map");
	printf("\nReading input map file from: %s", buf);
 	if ((map = (GbMap *)malloc(sizeof(GbMap))) == NULL) {
		fprintf(stdout, "Memory allocation failure for map\n");
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
	if ((status = ReadMapFile(buf, map)) != CE_OK) {
		fprintf(stdout, "Error %d from ReadMapFile(%s)\n", status, buf);
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
	
	// Read the input target density file.
	sprintf(buf, "%s%s", path, "PictoIN.den");
	printf("\nReading input target density file from: %s", buf);
	if ((status = ReadTargetFile(buf, &target)) != CE_OK) {
		fprintf(stdout, "Error %d from ReadTargetFile(%s)\n", status, buf);
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
	
	// Read the input measurement file.
	sprintf(buf, "%s%s", path, "PictoIN.mes");
	printf("\nReading input measurement file from: %s", buf);
 	if ((measDens = (GbColorInt *)malloc((GB_MAX_TP_STEPS * 4) * sizeof(GbColorInt))) == NULL) {
		fprintf(stdout, "memory allocation failure\n");
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
 	if ((status = ReadMeasDensFile(buf, (GB_MAX_TP_STEPS * 4), measDens, &nMeas)) != CE_OK) {
		fprintf(stdout, "Error %d from ReadMeasDensFile(%s)\n", status, buf);
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}

	// Pre-flight the densities.
	printf("\nPre-flighting densities...\n");
 	GbExamineDensities(nMeas, measDens, &target, &report);
	PrintReport(&report);
	
	// Refine the map.
	if ((status = GbRefineGrayMap(map, map, nMeas, measDens, &target)) != CE_OK) {
		fprintf(stdout, "Error %d from RefineGrayMap\n", status);
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
	
	// Write the output map file.
	sprintf(buf, "%s%s", path, "PictoOUT.map");
	printf("\nWriting output map file to: %s", buf);
	if ((status = WriteMapFile(buf, map)) != CE_OK) {
		fprintf(stdout, "Error %d from WriteMapFile(%s)\n", status, buf);
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
	
	// Write the output target density file.
	sprintf(buf, "%s%s", path, "PictoOUT.den");
	printf("\nWriting output target density file to: %s", buf);
	if ((status = WriteTargetFile(buf, &target)) != CE_OK) {
		fprintf(stdout, "Error %d from WriteTargetFile(%s)\n", status, buf);
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
	
	// Clean up.
	printf("\nExiting...\n\n");
 	ErrorCleanup(&target, map, measDens);
//	ExitNow(0);
}


static void ErrorCleanup(GbTarget *target, GbMap *map, GbColorInt *measDens)
{
	if (target->digital != NULL)
		free(target->digital);
	if (target->density != NULL)
		free(target->density);
	if (map != NULL)
		free(map);
	if (measDens != NULL)
		free(measDens);
}


static void ExitNow(int exitStatus)
{
	char buf[10];
	
	buf[0] = 'x';
	fflush(stdin);
	while ((buf[0] != 'Q') && (buf[0] != 'q')) {
		printf("Enter Q to Quit: ");
		scanf("%s", buf);
	}
	exit(exitStatus);
}	
