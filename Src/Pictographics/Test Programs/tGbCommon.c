/*
 *	tCommon.c - Common routines for testing GrayBalLib software.
 */

#include "tGrayBalLib.h"



void
GetNonCommentLine(FILE *fd, char *buf)
{
	buf[0] = '\0';
	do {
		fgets(buf, BUF_SIZE, fd);
	} while (buf[0] == '#');
}


void
PrintReport(GbDensReport *report)
{
	if (report->dMaxError >= 0)
		fprintf(stdout, "Dmax error = %5.3f\n", (Flt)report->dMaxError * 0.001);
	if (report->dMinError >= 0)
		fprintf(stdout, "Dmin error = %5.3f\n", (Flt)report->dMinError * 0.001);
	if (report->grayBalRMS >= 0)
		fprintf(stdout, "Gray balance: RMS = %5.3f, max = %5.3f at patch %d\n", (Flt)report->grayBalRMS * 0.001, report->grayBalMaxErr * 0.001, report->grayBalStep + 1);
	if (report->badMeasMaxErr >= 0)
		fprintf(stdout, "Largest stray = %5.3f at patch %d\n", (Flt)report->badMeasMaxErr * 0.001, report->badMeasStep + 1);
}


int
ReadMapFile(char *mapFileName, GbMap *map)
{
	char buf[128];
	FILE *fd;
	int i, c;


	if ((mapFileName == NULL) || (map == NULL))
		return(CE_ARGERROR);

	/* read the file */
	if ((fd = fopen(mapFileName, "r")) == NULL)
		return(CE_OPENERROR);

	if (fgets(buf, sizeof(buf), fd) == NULL) {
		fclose(fd);
		return(CE_READERROR);
	}
	if (strncmp(buf, "Version = 1", 11) != 0) {
		fclose(fd);
		return(CE_FORMATERROR);
	}

	if (fgets(buf, sizeof(buf), fd) == NULL) {
		fclose(fd);
		return(CE_READERROR);
	}
	if (strncmp(buf, "Type = LUT", 10) != 0) {
		fclose(fd);
		return(CE_FORMATERROR);
	}

	if (fgets(buf, sizeof(buf), fd) == NULL) {
		fclose(fd);
		return(CE_READERROR);
	}
	if (strncmp(buf, "Rows = 256", 10) != 0) {
		fclose(fd);
		return(CE_FORMATERROR);
	}

	if (fgets(buf, sizeof(buf), fd) == NULL) {
		fclose(fd);
		return(CE_READERROR);
	}
	if (strncmp(buf, "Cols = 3", 8) != 0) {
		fclose(fd);
		return(CE_FORMATERROR);
	}

	if (fgets(buf, sizeof(buf), fd) == NULL) {
		fclose(fd);
		return(CE_READERROR);
	}
	if (strncmp(buf, "UseXValues = 0", 14) != 0) {
		fclose(fd);
		return(CE_FORMATERROR);
	}

	for (i = 0; i < 256; i++) {
		if (fgets(buf, sizeof(buf), fd) == NULL) {
			fclose(fd);
			return(CE_READERROR);
		}
		if (sscanf(buf, "%f, %f, %f", &map->lut[0][i], &map->lut[1][i], &map->lut[2][i]) != 3) {
			fclose(fd);
			return(CE_FORMATERROR);
		}
		for (c = 0; c < 3; c++) {
			if ((map->lut[c][i] < 0.0) || (map->lut[c][i] > 1.0)) {
				fclose(fd);
				return(CE_FORMATERROR);
			}
		}
	}

	fclose(fd);

	return(CE_OK);
}


int
ReadMeasDensFile(char *inMeasFile, int nMeasMax, GbColorInt *measDens, int *nMeas)
{
	char buf[BUF_SIZE], c;
	int i;
	FILE *fd;
	Flt rDen, gDen, bDen;


	if (nMeasMax < 1)
		return(CE_ARGERROR);
	if ((fd = fopen(inMeasFile, "r")) == NULL)
		return(CE_OPENERROR);

	GetNonCommentLine(fd, buf);

	*nMeas = 0;
	while (1) {
		if (fgets(buf, sizeof(buf), fd) == NULL)
			break;
		if (sscanf(buf, "%c%d %f %f %f", &c, &i, &rDen, &gDen, &bDen) != 5)
			break;
		if (*nMeas == nMeasMax)
			return(CE_FORMATERROR);
		measDens[*nMeas][0] = (int)(rDen * 1000.0 + 0.5);
		measDens[*nMeas][1] = (int)(gDen * 1000.0 + 0.5);
		measDens[*nMeas][2] = (int)(bDen * 1000.0 + 0.5);
		*nMeas += 1;
	}
	fclose(fd);
	return(CE_OK);
}


int
ReadTargetFile(char *targetFileName, GbTarget *target)
{
	char buf[BUF_SIZE];
	FILE *fd;
	Flt rDen, gDen, bDen;
	int i, j;


	if ((fd = fopen(targetFileName, "r")) == NULL)
		return(CE_OPENERROR);
	target->modMethod = MY_TARGET_MOD_METHOD;
	if (target->modMethod == GB_TARGET_MOD_METHOD_CLAMP)
		fprintf(stdout, "target mod method = clamp\n");
	else if (target->modMethod == GB_TARGET_MOD_METHOD_LINEAR)
		fprintf(stdout, "target mod method = linear\n");
	else if (target->modMethod == GB_TARGET_MOD_METHOD_LINEAR_MINMAX)
		fprintf(stdout, "target mod method = linear min/max\n");
	fflush(stdout);
	GetNonCommentLine(fd, buf);
	GetNonCommentLine(fd, buf);
	GetNonCommentLine(fd, buf);
	GetNonCommentLine(fd, buf);
	GetNonCommentLine(fd, buf);
	GetNonCommentLine(fd, buf);
	if (sscanf(buf, "%d", &target->nTargetSamples) != 1) {
		fclose(fd);
		return(CE_FORMATERROR);
	}
	if ((target->digital = (GbColorInt *)malloc(target->nTargetSamples * sizeof(GbColorInt))) == NULL) {
		fclose(fd);
		return(CE_MALLOCERROR);
	}
	if ((target->density = (GbColorInt *)malloc(target->nTargetSamples * sizeof(GbColorInt))) == NULL) {
		free(target->digital);
		target->digital = NULL;
		fclose(fd);
		return(CE_MALLOCERROR);
	}
	for (i = 0; i < target->nTargetSamples; i++) {
		GetNonCommentLine(fd, buf);
		if (sscanf(buf, "%d %d %d %d %f %f %f",
			&j,
			&target->digital[i][0], &target->digital[i][1], &target->digital[i][2],
			&rDen, &gDen, &bDen) != 7) {
			free(target->digital);
			target->digital = NULL;
			free(target->density);
			target->density = NULL;
			fclose(fd);
			return(CE_MALLOCERROR);
		}
		target->density[i][0] = (int)(rDen * 1000.0 + 0.5);
		target->density[i][1] = (int)(gDen * 1000.0 + 0.5);
		target->density[i][2] = (int)(bDen * 1000.0 + 0.5);
	}
	fclose(fd);
	return(CE_OK);
}


int
WriteMapFile(char *mapFileName, GbMap *map)
{
	FILE *fd;
	int i;


	if ((mapFileName == NULL) || (map == NULL))
		return(CE_ARGERROR);

	if ((fd = fopen(mapFileName, "w")) == NULL)
		return(CE_OPENERROR);

	fprintf(fd, "Version = 1\n");
	fprintf(fd, "Type = LUT\n");
	fprintf(fd, "Rows = 256\n");
	fprintf(fd, "Cols = 3\n");
	fprintf(fd, "UseXValues = 0\n");

	for (i = 0; i < 256; i++)
		fprintf(fd, "%8.6f, %8.6f, %8.6f\n", map->lut[0][i], map->lut[1][i], map->lut[2][i]);

	fclose(fd);

	return(CE_OK);
}


int
WriteTargetFile(char *targetFileName, GbTarget *target)
{
	FILE *fd;
	int i;


	if ((fd = fopen(targetFileName, "w")) == NULL)
		return(CE_OPENERROR);
	fprintf(fd, "Line 1\n0\n0\n255\n4500\n%d\n", target->nTargetSamples);
	for (i = 0; i < target->nTargetSamples; i++) {
		fprintf(fd, "%3d %3d %3d %3d %5.3f %5.3f %5.3f\n",
			i + 1,
			target->digital[i][0], target->digital[i][1], target->digital[i][2],
			(Flt)target->density[i][0] * 0.001, (Flt)target->density[i][1] * 0.001, (Flt)target->density[i][2] * 0.001);
	}
	fclose(fd);
	return(CE_OK);
}
