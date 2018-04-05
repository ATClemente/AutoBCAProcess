/**********************************
**  Author: Anthony Clemente     **
**    Date: 10 February 2017     **
***********************************/


#include <fstream>
#include <iostream>
#include <string>

//Library for working with Excel
#include "libxl.h"

//Library for working with Excel
using namespace libxl;



/**********************************
**  DECLARE NECESSARY CONSTANTS  **
***********************************/

//Number of bytes per a line with a value in notepad
//7 bytes because format is X.XXX plus the new line which is 2 bytes
const int SIZE_OF_VALUE_LINE = 7;

//Set of 8 standards plus a blank
//This is often multiplied by two since everything is done in duplicate
const int NUM_STANDARDS = 9;

//To skip a standard set in the file it is 63 bytes (9 lines for standards @ 7 bytes per line)
const int STANDARD_SET_BYTES = NUM_STANDARDS * SIZE_OF_VALUE_LINE;

//The line in the notepad output "Second set:" is equal to 13 bytes
const int SIZE_OF_SECONDSET_LINE = 13;

//The two spacing blank lines are equal to 2 bytes each
const int SIZE_OF_BLANK_LINE = 2;


//Marks where the standards start.  Currently it is row 35 but since
//the first row is considered row 0, 1 is subtracted from this.
//If this ever needs to change, only change the "35"
const int BLANK_ROW_START = 35 - 1;

//Same idea but for the column start
const int BLANK_COL_START = 3 - 1;

//Same thing for where the experimental data is entered in the sheet
const int EXPERIMENTAL_ROW_START = 68 - 1;
const int EXPERIMENTAL_COL_START =  4 - 1;



void writeStandardSet(std::ifstream &inputFile, Sheet* &outputSheet, int numWells);

void writeExperimentalSet(std::ifstream &inputFile, Sheet* &outputSheet, int numWells);


int main()
{

	std::string inputFileName;

	std::cout << "Enter a text file: ";
	std::getline(std::cin, inputFileName);

	int numWells;
	std::cout << "Enter total wells used: ";
	std::cin >> numWells;

	std::ifstream inputFile;

	inputFile.open(inputFileName);

	//"Create" an Excel (in this case .xlsx, .xls is also possible) book to work with
	//for .xls use "xlCreateBook()"
	if (inputFile)
	{
		Book* outputExcelFile = xlCreateXMLBook();

		if (outputExcelFile)
		{
			//Loads an existing file by name
			if (outputExcelFile->load(L"TemplateBSA.xlsx"))
			{
				//Accesses the first sheet in the book
				Sheet* sheet = outputExcelFile->getSheet(0);
				if (sheet)
				{
					//Call the functions required to do the reading and writing of data
					std::cout << "Writing standards to file..." << std::endl;
					writeStandardSet(inputFile, sheet, numWells);
					std::cout << "Writing experimentals to file..." << std::endl;
					writeExperimentalSet(inputFile, sheet, numWells);
					std::cout << "Done!" << std::endl;
				}
				else
				{
					std::cout << "Cannot find that sheet in specified file." << std::endl;
				}
				outputExcelFile->save(L"TemplateBSA.xlsx");
			}
			else
			{
				std::cout << "Cannot find the internally specified excel file." 
					<<std::endl<< "Excel file must be named: \"TemplateBSA\" and be in the .xlsx format" << std::endl;
			}

			outputExcelFile->release();
		}
		else
		{
			std::cout << "Could not create book? or other error that prevented main if statement execution" << std::endl;
		}
	}
	else
	{
		std::cout << "Cannot find that text file." << std::endl;
	}


	inputFile.close();


	return 0;
}

/*************************************************
**  Function used for writing the data for the  **
**  standards into the sheet. Both the input    **
**  text file and the output excel file are     **   
**  passed by reference.                        ** 
**************************************************/

void writeStandardSet(std::ifstream &inputFile, Sheet* &outputSheet, int numWells)
{
	//Double to hold the current value being read from the sheet.
	double currentValue;
	int loopCount = 0;
	int rowNum = BLANK_ROW_START;
	int colNum = BLANK_COL_START;

	//Starts at beginning of file and starts reading values.
	//Loop checks to make sure that we are reading a value and have not gone past the standards
	//Loop checks counter first because checking "inputFile >> currentValue" causes the program
	//to attempt to read a value before evaluating. This can cause errors if whitespace is
	//attempted to be read as a double.  It will evaluate to false, stopping the loop, but
	//the filestream will now have an error. By checking the loop counter first, this avoids the
	//problem for some later loops where this is an issue.
	while (loopCount < 9 && (inputFile >> currentValue))
	{
		//From excel library, writes the data desired to the cell corresponding
		//to the given row and colmun
		outputSheet->writeNum(rowNum, colNum, currentValue);

		//Standard run from highest to lowest upward in the excel sheet,
		//so row num is decremented after each run.
		rowNum--;
		loopCount++;
	}


	//Since everything is duplicate, we need to run again to fill the other column of 
	//standards. Reset all appropriate variables while adjusting the column start point
	loopCount = 0;
	rowNum = BLANK_ROW_START;
	colNum = BLANK_COL_START + 1;


	//Uses the file stream "seekg" function (offset from beginning of file) to make sure the next data read is from the appropriate spot.
	//The calculated number of bytes moves to the set of standards that have not been read yet.
	//Check the contants declaration at the top of this file for the context of these numbers.
	int seekSpot = STANDARD_SET_BYTES + (((numWells - (NUM_STANDARDS*2)) / 2) * SIZE_OF_VALUE_LINE) + SIZE_OF_SECONDSET_LINE + (SIZE_OF_BLANK_LINE * 2);

	inputFile.seekg(seekSpot, std::ios::beg);


	//Repeats the same process but now reading values from a different spot and writing to a
	//different column
	while (loopCount < 9 && (inputFile >> currentValue))
	{
		outputSheet->writeNum(rowNum, colNum, currentValue);
		rowNum--;
		loopCount++;
	}

}



/*************************************************
**  Function used for writing the data for the  **
**  experimental points into the sheet. Both    **
**	the input text file and the output excel    **
**	file are passed by reference. This function **
**	works almost identically to the function    **
**	for standards but uses different values to  **
**	read and write data at the appropriate      **
**	locations.                                  **
**************************************************/
void writeExperimentalSet(std::ifstream &inputFile, Sheet* &outputSheet, int numWells)
{
	double currentValue;
	int loopCount = 0;

	//Different values representing the appropriate location for entering
	//experimentals into the target sheet
	int rowNum = EXPERIMENTAL_ROW_START;
	int colNum = EXPERIMENTAL_COL_START;

	//Skips the standards
	inputFile.seekg(STANDARD_SET_BYTES, std::ios::beg);

	//Same while loop but note that the rowNum is incremented in this case not decremented
	//Writes the first column of experimentals
	while (loopCount < ((numWells - (NUM_STANDARDS * 2)) / 2) && (inputFile >> currentValue))
	{
		outputSheet->writeNum(rowNum, colNum, currentValue);
		rowNum++;
		loopCount++;
	}



	//Reset for second looping
	loopCount = 0;
	rowNum = EXPERIMENTAL_ROW_START;
	colNum = EXPERIMENTAL_COL_START + 1;

	//Calculated number of bytes to skip in order to seek to the proper location
	int seekSpot = (STANDARD_SET_BYTES * 2) + (((numWells - (NUM_STANDARDS * 2)) / 2) * SIZE_OF_VALUE_LINE) + SIZE_OF_SECONDSET_LINE + (SIZE_OF_BLANK_LINE * 2);

	inputFile.seekg(seekSpot, std::ios::beg);

	//Write the second column of experimentals
	while (loopCount < ((numWells - (NUM_STANDARDS * 2)) / 2) && (inputFile >> currentValue) )
	{
		outputSheet->writeNum(rowNum, colNum, currentValue);
		rowNum++;
		loopCount++;
	}

}

