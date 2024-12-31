// PictoLUT.cpp : Defines the entry point for the console application.

#include <io.h>
#include <conio.h>
#include <stdio.h>
#include <stdlib.h>
// #include "stdafx.h"
#include "GrayBalLib.h"
#include <vcl.h>
#pragma hdrstop

__declspec(dllimport) int
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


GbMap inGMap;
GbMap outGMap;

GbColorInt measDens;
GbTarget target;
int nMeas;

int main(int argc, char* argv[])
{
 	int Result = 0;

	printf("\nRedmer Controls Wrapper for Pictographics GrayBalLib DLL version 1.00\n");


	nMeas = 1;
	Result = GbRefineGrayMap(&inGMap, &outGMap, nMeas, &measDens, &target);

	printf("\nGbRefineGrayMap returned %d\n", Result);


        getch();

	return 0;
}
