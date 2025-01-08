from my_package import extract_quarter_raw

dir = r"/root_path_to_working_folder"
file = "PPE&DWE_SOUTH_DR_2024.xlsx"
save = "4th QUARTER ACCOMPLISHMENT REPORT 2024.xlsx"
sheets = ["OCT","NOV","DEC"]

extract_quarter_raw(dir, file, save, sheets)