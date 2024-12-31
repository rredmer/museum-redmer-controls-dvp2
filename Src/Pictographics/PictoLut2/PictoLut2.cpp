// PictoLut2.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "GrayBalLib.h"



GbMap inGMap;
GbMap outGMAP;
GbColorInt measDens;
GbTarget target;
int nMeas;

int main(int argc, char* argv[])
{
	int Result = 0;

	printf("\nRedmer Controls Wrapper for Pictographics GrayBalLib DLL version 1.00\n");


	nMeas = 1;
	Result = GbRefineGrayMap(&inGMap, &outGMAP, nMeas, &measDens, &target);
	
	printf("\nGbRefineGrayMap returned %d\n", Result);

	return 0;
}

