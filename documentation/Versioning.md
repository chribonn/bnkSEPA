1.00.00	First Version of this Spreadsheet addressing BoV needs 

1.10.00
  * Added Client Payee Data Tab. This tab allows client to add a sorted database of their clients. Using the VLOOKUP function, the Credit	Instruction Record screen can be populated with the payee data rather than having to be typed in. 
  * Renamed the tab Control Data to Control Data (Hidden). 
  * In the Control Data (Hidden) tab, the records from Payment Information Record are copied into a contiguous list and a Pivot table summarising the payment batch is shown in the Control tab. This allows the poster to generate a payment breakdown."  

1.10.10	
  * PIR: PstlAdr:AdrLine - this field is mandatory as instructed by BoV

1.20.00	
  * Increased the amount of rows in CIR from 32 -> 51 
    * Adjusted Control Tab to reflect new entries 
    * Control Tab Columns Formatting is allowed 

1.30.00	
  * Payment Information Record - PstlAdr:AdrLine  
    * Do not write output if when either cell is empty.  Original behaviour was that if Address Line 1 was not blank both lines are written.

  * Implemented Sub **ClearSheet()** in WorkSheet *Credit Instruction Record* that goes through all the rows and blanks them.
    *  Code is not linked to a button event. 

1.35.00	
  * The passwords that are used to open the zipped file and to encrypt the SCTE file can now be passed as parameters. If they are absent the code will default to entries in **secrets.py**.
  * The name of the xlsm file that is within the zipped archive is a parameter rather than being hard coded into the code.  The default name is **BnkSEPA.xlsm**. This adds another layer of security in an embedded solution.
  * The VBA code to update the PivotTable in the Control worksheet has been updated. This is because the PivotTable was not always updating. The behaviour is described in https://techcommunity.microsoft.com/t5/excel/pivot-table-won-t-refresh-after-data-refresh/m-p/2937032 

1.37.00	
  * Applied changes defined in the SEPA Credit Transfers file layout document dated 2022-03
  * Added the button call the macro that blanked the CIR Worksheet

1.38.01	
  * Applied changes defined in the SEPA Credit Transfers file layout document dated 2023-09
  * Code review and improvement
  * Minor bug fixes
  * Updated requirements.txt
  * Tested (and env upgrade) to Pyhon 3.11

2.00.000	
  * The generated **SCT** file is not transformed into a password-protected **SCTE** file.  All logic related to this function has been removed.
  * Minor documentation corrections

## Help the project

[Click here](/documentation/HelpbnkSEPA.md) to read how you can help the project.

