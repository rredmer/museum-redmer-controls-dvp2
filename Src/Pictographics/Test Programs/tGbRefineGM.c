/*
 *	tGbRefineGM.c - test the GbRefineGM.c software.
 */

#include "tGrayBalLib.h"

#define	GB_MAX_TP_STEPS		256	/* maximum number of steps allowed in gray balance test pattern */


static void	ErrorCleanup(GbTarget *target, GbMap *map, GbColorInt *measDens);
static void	ExitNow(int exitStatus);


void
main(int argc, char **argv)
{
	char buf[256];
	GbColorInt *measDens;
	GbDensReport report;
	int status, nMeas;
	GbTarget target;
	GbMap *map;


	printf("tGbRefineGM\n");
	
/*
 *	Initialize.
 */
	map = NULL;
	target.digital = NULL;
	target.density = NULL;
	measDens = NULL;
	
/*
 *	Read the input map file.
 */
	printf("Enter inMapFile: ");
	scanf("%s", buf);
	printf("\n");

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
	
/*
 *	Read the input target density file.
 */
	printf("Enter inTargetFile: ");
	scanf("%s", buf);
	printf("\n");

	if ((status = ReadTargetFile(buf, &target)) != CE_OK) {
		fprintf(stdout, "Error %d from ReadTargetFile(%s)\n", status, buf);
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
	
/*
 *	Read the input measurement file.
 */
	printf("Enter inMeasFile: ");
	scanf("%s", buf);
	printf("\n");

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
	
/*
 *	Pre-flight the densities.
 */
	GbExamineDensities(nMeas, measDens, &target, &report);
	PrintReport(&report);
	
/*
 *	Refine the map.
 */
	if ((status = GbRefineGrayMap(map, map, nMeas, measDens, &target)) != CE_OK) {
		fprintf(stdout, "Error %d from RefineGrayMap\n", status);
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
	
/*
 *	Write the output map file.
 */
	printf("Enter outMapFile: ");
	scanf("%s", buf);
	printf("\n");

	if ((status = WriteMapFile(buf, map)) != CE_OK) {
		fprintf(stdout, "Error %d from WriteMapFile(%s)\n", status, buf);
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
	
/*
 *	Write the output target density file.
 */
	printf("Enter outTargetFile: ");
	scanf("%s", buf);
	printf("\n");

	if ((status = WriteTargetFile(buf, &target)) != CE_OK) {
		fprintf(stdout, "Error %d from WriteTargetFile(%s)\n", status, buf);
		ErrorCleanup(&target, map, measDens);
		ExitNow(-1);
	}
	
/*
 *	Clean up.
 */
	ErrorCleanup(&target, map, measDens);
	ExitNow(0);	/* success */
}


static void
ErrorCleanup(GbTarget *target, GbMap *map, GbColorInt *measDens)
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


static void
ExitNow(int exitStatus)
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
