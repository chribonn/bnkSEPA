# Embedding the solution into a portal

The design of this solution was built so that it could be integrated into a larger solution that  part of a client facing portal with this being one of the options.

![](../images/overview001.png)

1. The user would log into a portal using their credentials (Username, Password, 2FA)
2. The authenticated user uploads the password-protected zip to the portal. *The contents of the file cannot be read while in transit.*
3. The system would extract the contents of the zip archive using the password fetched from the secure database (**zippass**).
4. The extracted XLSM file name would be validated against what is expected (**xlfile**). *This constitutes an additional layer of security*.
5. The XL file is transformed into the SCT file format.
6. The SCT file is converted into SCTE using the financial-institution agreed password (**bankSCTE**).
7. Once the SCTE file is ready, the solution would:
    - (async) transmit the file to the financial institution directly 
    - (sync) allow the user to download the file to perform the transmission part themselves.
8. Code cleanup.
    
The parameter **bankSCTE** could be determined without the knowledge of the person who uploads the file. *This results in a solution in which the end-2-end process is not know to a single person.
