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
  * The SCTE password that was embedded in the XL file has been removed.
 
