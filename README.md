Welcome to pandabear...
---
This project was started as a tool to help the scheduling workflow of a supervisor.

How it works:
  **pandabear** takes an existing excel scheduling document and converts a specified page into a format that Microsoft Teams will accept as an import.
  Upon running it initially opens as an interactive UI with two buttons: Upload File, and Employee Emails.
    - Upload file will be a guided file dialog selection of the _.xlsx_ file that the user would like to convert, and, upon selection, a scroll-able field will prompt for specific sheet selection.
    - Employee Emails on the other hand will, once clicked, open an interactive display that directly manipulates a local _.json_ file that ties employee names with their Microsoft emails.
  After the converted file appears in the file explorer, the user will then be able to select it from the Microsoft Teams/Shifts upload button.

How to use it (_assuming it has been packaged into a single .exe appropriately_):
  - After running the .exe you will select an .xlsx file with the appropriate layout (provided in repo's folder)
  - Post-selection a scroll-able field will appear prompting for a specific sheet selection
  - Once the doc and sheet has been selected there will be a save-as prompt for deciding the new file's name and where to save it
  - Abrakadabra.. there is now a properly formatted excel document ready to be imported into Microsoft Teams/Shifts
  - The Employee Email button is for correlating how the employee's name is recorded in the original excel document. For accurate use all employees must be correlated with an email in this field   

  Disclaimer...
    This projects syntax was coded entirely with ai. With that being said I did beforehand internalize the exact architecture of what and how I wanted the program to do what it does.
      It was a very fascinating experience and humbling in ways that I did not expect. Without the crutch of ai as a next-level interpreter I would not have completed this project in a matter of hours.
      Regardless of the methods, it will now fix a very time-consuming issue that a supervisor was facing.

  In closing, I feel the need to speak my mind. I used to think that coding and engineering solutions was satisfying due to me being the one to make every nuance of it, including the code itself.
    There is some sort of pride in bragging about how long you have spent creating a solution or how long you have spent on an issue at hand. However, as cliche as it may seem, technology is evolving
    at a rapid rate. I believe that ai should not replace the fundamental understanding of code, syntax or the core of HOW a project works especially, but ai seems to be the upgraded code interpreter. 
    Instead of inputting a coding language, we can simply code with our human language. This is next level and I plan on keeping up, but admittedly I will have to put that part of my ego aside.
