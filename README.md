# UOM 1st Semester Rank Generator (v4.00)

## Requirements:

-   Java 8 or higher
-   Python 3.8 or higher

## Instructions:

1.  Add the PDF file containing the provincial results for each module into the `data/provisional results/` directory in the root folder.
2.  Rename the pdf files with `module code`
3.  Add the `mpr_student_details.json` file in the `data/` folder.
4.  Configure the `config.ini` file in the `root` folder.
5.  Change the course according to the reuired file (i.e `mpr`, `tm`, `em`, `tmle`)
6.  Run the Python program.
7.  `Results Analysis.xlsx` and `Results Analysis (extended).xlsx` files will be generated which contains the following info,
    `Rank` - Rank is given considering GPA on a **4.0 scale**.
    `Batch Rank` - Rank is given considering GPA on **4.0 scale** and **4.2 scale**.
    `Current SGPA` - Current Grade Point Average for the semester for the modules that results are released.
    `Maximum Possible SGPA` - Maximum Grade Point Average for the semester that the student can achieve.
8.  Sorting is done by the _lexicographical ordering_ of following factors in descending order. (If you want to change the order of priority of modules, change the order in `config.ini`)
    1. Rank
    2. Batch Rank
    3. MA1014 Grade
    4. CS1033 Grade
    5. EE1040 Grade
    6. ME1033 Grade
    7. CE1023 Grade
    8. MT1023 Grade
    9. Index (Ascending)
