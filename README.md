# PsychInterviewPacksAutomata

What it does:
- Creates each applicants interview pack folder 
- Renames the Referee reports downloaded from HODSPA and moves into folder 
- Moves the Coversheets into the folder 

Instructions:
1. Place the script and the Excel sheet containing the FirstName in column A, GivenName in column B, and Student ID in column C
2. Change the name of the Excel sheet in the readmatrix() function to whatever you are using
3. Download all the relevant Referee reports (.pdf) and create the Coversheets (.docx) into the folder containing the script
4. Ensure that the FirstName and GivenName are present on the .pdf and .docx files. The code will match these to the foldername
5. Ensure that no students have the same FirstName and GivenName, would recommend doing in batches if there are
6. Run the entire script and the code will handle the rest :)
7. The code will skip over some students with complicated names or special characters. In this case you'll have to manually rename and move them after you've run the code
8. Drag the interview packs to upload it onto sharepoint
