/*
 *	tRefineES.c - test the RefineES.c software.
 */

#include "tGrayBalLib.h"

#define	MAX_MEAS	30

static void	ExitNow(int exitStatus);
static int	FindTargetDmax(GbTarget *target);
static int	FindTargetDmin(GbTarget *target);


void
main(int argc, char **argv)
{
	char buf[256];
	GbColorInt inES, outES, stepSize, targetDens, *measDens;
	GbDensReport report;
	Flt eMin, eMax;
	int status, nMeas, c, iMin, iMax;
	GbTarget target;

	printf("tRefineEs\nEnter targetFile: ");
	scanf("%s", buf);
	printf("\n");
	if ((status = ReadTargetFile(buf, &target)) != CE_OK) {
		fprintf(stdout, "Error %d from ReadTargetFile %s\n", status, buf);
		ExitNow(-1);
	}
	printf("Enter expR expG expB: ");
	scanf("%d %d %d", &inES[0], &inES[1], &inES[2]);
	printf("\n");
	printf("Enter dExpR dExpG dExpB: ");
	scanf("%d %d %d", &stepSize[0], &stepSize[1], &stepSize[2]);
	printf("\n");
	printf("Enter inMeasFile: ");
	scanf("%s", buf);
	printf("\n");

	target.modMethod = GB_TARGET_MOD_METHOD_CLAMP;
/*
 *	Read the input measurement file.
 */
	if ((measDens = malloc(MAX_MEAS * sizeof(GbColorInt))) == NULL) {
		free(target.digital);
		free(target.density);
		fprintf(stdout, "memory allocation failure.\n");
		ExitNow(-1);
	}
	if ((status = ReadMeasDensFile(buf, MAX_MEAS, measDens, &nMeas)) != CE_OK) {
		free(target.digital);
		free(target.density);
		free(measDens);
		fprintf(stdout, "Error %d from ReadMeasDensFile\n", status);
		ExitNow(-1);
	}
	fprintf(stdout, "Number of measurements read = %d\n", nMeas);
	fprintf(stdout, "Actual dens = %5.3f %5.3f %5.3f\n", (Flt)measDens[0][0] * 0.001, (Flt)measDens[0][1] * 0.001, (Flt)measDens[0][2] * 0.001);
	fflush(stdout);
/*
 *	Pull the target densities (D max or D min depending on media polarity).
 */
	iMax = FindTargetDmax(&target);
	iMin = FindTargetDmin(&target);
	eMax = target.density[iMax][1] - measDens[0][1];
	eMax = (eMax < 0.0) ? -eMax : eMax;
	eMin = target.density[iMin][1] - measDens[0][1];
	eMin = (eMin < 0.0) ? -eMin : eMin;
	for (c = 0; c < 3; c++)
		targetDens[c] = (eMax < eMin) ? target.density[iMax][c] : target.density[iMin][c];

	fprintf(stdout, "Target dens = %5.3f %5.3f %5.3f\n", (Flt)targetDens[0] * 0.001, (Flt)targetDens[1] * 0.001, (Flt)targetDens[2] * 0.001);
	fflush(stdout);
/*
 *	Pre-flight the densities.
 */
	GbExamineDensities(nMeas, measDens, &target, &report);
	PrintReport(&report);
	free(target.digital);	/* done with these */
	free(target.density);
/*
 *	Refine the map.
 */
	if ((status = GbRefineExposureSet(inES, outES, stepSize, targetDens, nMeas, measDens)) != CE_OK) {
		fprintf(stdout, "Error %d from RefineExposureSet\n", status);
		free(measDens);
		ExitNow(-1);
	}
	free(measDens);		/* done with this */

	fprintf(stdout, "Exposures = %d %d %d\n", outES[0], outES[1], outES[2]);

	ExitNow(0);	/* success */
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


static int
FindTargetDmax(GbTarget *target)
{
	int i, maxSoFar, iSoFar;


	maxSoFar = -1;
	iSoFar = 0;
	for (i = 0; i < target->nTargetSamples; i++) {
		if (target->density[i][1] > maxSoFar) {
			maxSoFar = target->density[i][1];
			iSoFar = i;
		}
	}
	return(iSoFar);
}


static int
FindTargetDmin(GbTarget *target)
{
	int i, minSoFar, iSoFar;


	minSoFar = 10000;
	iSoFar = 0;
	for (i = 0; i < target->nTargetSamples; i++) {
		if (target->density[i][1] < minSoFar) {
			minSoFar = target->density[i][1];
			iSoFar = i;
		}
	}
	return(iSoFar);
}
