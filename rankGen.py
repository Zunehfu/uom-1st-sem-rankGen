# UOM rank generator for 1st semester (v4.00)
# ==========================================
# Written by Deneth Priyadarshana 

import tabula
import os.path
import xlsxwriter
import json
import configparser

print("-------------------------------------------------------------------")
print("UOM rank generator for 1st semester (v4.00)")
print("-------------------------------------------------------------------")
print("* Ignore the 'No module named jpype' warning if you see it.")
print("# Reading config.ini...")

config = configparser.ConfigParser()
config.optionxform = str
config.read("config.ini")    

if "VARIABLES" not in config:
    print("[ERROR]: 'VARIABLES' section not found.")
    exit()

COURSE = config["VARIABLES"]["course"]
RESULTS_PATH = config["VARIABLES"]["results_path"]
STUDENT_DETAILS_PATH = config["VARIABLES"]["student_details_path"]

MODULES = {}
TOTAL_CREDITS = 0
if "MODULE_INFO" in config:
    for module, credits in config.items("MODULE_INFO"):
        MODULES[module] = {}
        MODULES[module]["credits"] = int(credits)
        TOTAL_CREDITS += int(credits)
else:
    print("[ERROR]: 'MODULE_INFO' section not found.")
    exit()

GRADES = {}
if "GRADE_INFO" in config:
    for grade, credits in config.items("GRADE_INFO"):
        GRADES[grade] = float(credits)
else:
    print("[ERROR]:'GRADE_INFO' section not found.")
    exit()

# Loading student details
print(f"# Reading '{STUDENT_DETAILS_PATH}'...")
with open(STUDENT_DETAILS_PATH, "r") as f:
   data = json.load(f)

stud_details_dict = {entry["index"]: entry for entry in data[COURSE]}
print(stud_details_dict)
INDEX_START = int(list(stud_details_dict)[0])
INDEX_END = int(list(stud_details_dict)[-1])

def is_in_course(index):
    return  INDEX_START <= index <= INDEX_END

def grade_to_gpa(grade, virtual = True):
    if grade == "A+":
        if virtual:
            return GRADES["A+"]
        return GRADES["A"]
    return GRADES[grade]

res_dict = {}
available_modules = []
available_indexes = {}
available_total_credits = 0

for module in MODULES:
    path = RESULTS_PATH + module + ".pdf"
    if os.path.isfile(path):
        print(f"# Reading '{path}'...")
        available_modules.append(module)
        available_total_credits += MODULES[module]["credits"]
        
        grade_tables = tabula.read_pdf(path, pages="all", pandas_options={'header': None})  
        index_grade_tuples = []
        
        for tbl in grade_tables: 
            index_grade_tuples += list(zip(list(tbl[0][1:]), list(tbl[1][1:])))
        
        for tup in index_grade_tuples:
                if str(tup[0]) != "nan" and is_in_course(int(tup[0][:-1])):
                    idx = int(tup[0][:-1])
                    if idx not in res_dict: 
                        res_dict[idx] = {}
                    res_dict[idx][module] = tup[1]  
                    MODULES[module][tup[1]] = MODULES[module].get(tup[1], 0) + 1
    
print("# Prepairing table...")

# Handle missing people only for MPR (others may raise unexpected bugs)
if COURSE == "mpr":
    missing_indexes = list(set(range(INDEX_START, INDEX_END+1)) - set(res_dict.keys()))
    for missing_index in missing_indexes:
        res_dict[missing_index] = {}
        for module in available_modules:
            # Consider withdrawn
            res_dict[missing_index][module] = "W" 
            MODULES[module]["W"] = MODULES[module].get("W", 0) + 1

# Handle missing modules in available people
for key in res_dict:
    if len(res_dict[key]) != len(available_modules):
         for module in available_modules:
             if module not in res_dict[key]:
                 # Consider withdrawn
                 res_dict[key][module] = "W" 

# current GPA (4.0 scale)
def get_gpa(dict):
    cg_product = 0
    for module in available_modules:
        cg_product += MODULES[module]["credits"] * grade_to_gpa(dict[module], False)        
    return round(cg_product/available_total_credits, 2)

# current GPA (4.2 scale)
def get_virtual_gpa(dict):
    cg_product = 0
    for module in available_modules:
        cg_product += MODULES[module]["credits"] * grade_to_gpa(dict[module])
    return cg_product/available_total_credits

# Maximum possible GPA (4.0 scale)
def get_max_gpa(dict):
    max_cg_product, cg_product = 0, 0 
    for module in available_modules:
        cg_product += MODULES[module]["credits"] * grade_to_gpa(dict[module], False)
        max_cg_product +=  MODULES[module]["credits"] * 4
    return round((4.0 - ((max_cg_product - cg_product)/TOTAL_CREDITS)), 2)
    
for idx in res_dict:
    res_dict[idx]["gpa"], res_dict[idx]["vgpa"], res_dict[idx]["mgpa"] = get_gpa(res_dict[idx]), get_virtual_gpa(res_dict[idx]), get_max_gpa(res_dict[idx])


# Sorting logic
# ------------------------
# 1. Sort students by GPA on a 4.0 scale in descending order.
# 2. If GPAs are equal, sort by GPA on a 4.2 scale in descending order.
# 3. If both GPAs are equal, sort by harder subjects in descending order.
def sort_key(student):
    gpa = student[1]["gpa"]
    vgpa = student[1]["vgpa"]
    module_gpas = [grade_to_gpa(student[1][module]) for module in available_modules]
    return (gpa, vgpa, *module_gpas, -student[0])

# Sort the dictionary using the dynamic sort key
res_dict = dict(sorted(res_dict.items(), key=lambda student: sort_key(student), reverse=True))
# ------------------------


# Add rankings
# ------------------------
prev_gpa, rank_gap , rank = 0, 0, 1
prev_vgpa, brank_gap, brank = 0, 0, 1
for idx in res_dict:
    if (prev_gpa == res_dict[idx]["gpa"]):
        res_dict[idx]["rank"] = rank
        rank_gap += 1
        if(prev_vgpa == res_dict[idx]["vgpa"]):
            res_dict[idx]["brank"] = brank
            brank_gap += 1
        else:
            brank += brank_gap
            res_dict[idx]["brank"] = brank
            brank_gap = 1
            prev_vgpa = res_dict[idx]["vgpa"]
    else: 
        rank += rank_gap
        brank += brank_gap
        res_dict[idx]["rank"] = rank
        res_dict[idx]["brank"] = brank
        rank_gap = 1
        brank_gap = 1
        prev_gpa = res_dict[idx]["gpa"]
        prev_vgpa = res_dict[idx]["vgpa"]
# ------------------------


# Export to excel
print("# Exporting to '.xlsx' files...")
# ------------------------ file 1
n = len(available_modules)
tot = len(MODULES)

workbook = xlsxwriter.Workbook(f"{COURSE.upper()} - Result Analysis.xlsx")
worksheet = workbook.add_worksheet()

# Students table
worksheet.write(0, 0, "Rank")
worksheet.write(0, 1, "Index") 
for i in range(n):
    worksheet.write(0, i + 2, available_modules[i])
if n != tot: worksheet.write(0,  n + 2, "Current SGPA")
else: worksheet.write(0,  n + 2, "SGPA")
if n != tot: worksheet.write(0,  n + 3, "Maximum Possible SGPA")

for i, idx in enumerate(res_dict):
    worksheet.write(i + 1, 0, res_dict[idx]["rank"])
    worksheet.write(i + 1, 1, stud_details_dict[idx]["full_index"])
    for j in range(n):
        worksheet.write(i + 1, j + 2, res_dict[idx][available_modules[j]])
    worksheet.write(i + 1,  n + 2,  res_dict[idx]["gpa"])
    if n != tot: worksheet.write(i + 1,  n + 3, res_dict[idx]["mgpa"])

# Grade Analysis table
k = n - 1
if n != tot: k = n

totals = [sum(list(MODULES[available_modules[i]].values())[1:]) for i in range(n)]

for i in range(n):
    worksheet.write(0, k + 7 + i, available_modules[i])

for m, grade in enumerate(GRADES):
    worksheet.write(m + 1, k + 6, grade)
    for i in range(n):
        v = MODULES[available_modules[i]].get(grade, 0)
        worksheet.write(m + 1, k + 7 + i, f"{v}({(v/totals[i])*100:.1f}%)")


workbook.close()

# ------------------------ file 2
workbook = xlsxwriter.Workbook(f"{COURSE.upper()} - Result Analysis (extended).xlsx")
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, "Rank")
worksheet.write(0, 1, "Index") 
worksheet.write(0, 2, "Name") 
worksheet.write(0, 3, "Group") 

for i in range(n):
    worksheet.write(0, i + 4, available_modules[i])
if n != tot: worksheet.write(0,  n + 4, "Current SGPA")
else: worksheet.write(0,  n + 4, "SGPA")
if n != tot: worksheet.write(0,  n + 5, "Maximum Possible SGPA")
if n != tot: worksheet.write(0,  n + 6, "Batch Rank (4.2 GPA scale)")
else: worksheet.write(0,  n + 5, "Batch Rank (4.2 GPA scale)")

for i, idx in enumerate(res_dict):
    worksheet.write(i + 1, 0, res_dict[idx]["rank"])
    worksheet.write(i + 1, 1, stud_details_dict[idx]["full_index"])
    worksheet.write(i + 1, 2, stud_details_dict[idx]["name"])
    worksheet.write(i + 1, 3, stud_details_dict[idx]["group"])
    for j in range(n):
        worksheet.write(i + 1, j + 4, res_dict[idx][available_modules[j]])
    worksheet.write(i + 1,  n + 4,  res_dict[idx]["gpa"])
    if n != tot: worksheet.write(i + 1,  n + 5, res_dict[idx]["mgpa"])
    if n != tot: worksheet.write(i + 1,  n + 6, res_dict[idx]["brank"])
    else: worksheet.write(i + 1,  n + 5, res_dict[idx]["brank"])

for i in range(n):
    worksheet.write(0, k + 10 + i, available_modules[i])

for m, grade in enumerate(GRADES):
    worksheet.write(m + 1, k + 9, grade)
    for i in range(n):
        v = MODULES[available_modules[i]].get(grade, 0)
        worksheet.write(m + 1, k + 10 + i, f"{v}({(v/totals[i])*100:.1f}%)")

workbook.close()
# ------------------------
    
print("Finished successfully!")
print("-------------------------------------------------------------------")