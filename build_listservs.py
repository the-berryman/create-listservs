import openpyxl
import shutil
import os
import datetime

def build_queries(LIST):
       
    with open (LIST, "w") as file:
        import openpyxl
        wb = openpyxl.load_workbook(filename='DynamicListCreation.xlsx', read_only=False)
        sheet = wb.get_sheet_by_name('Sheet1')
        for row in range(2, sheet.max_row + 1):
        
                DEPT_ID = sheet['A' + str(row)].value   
                DEPT_NAME = sheet['B' + str(row)].value 
                FAC_LIST = sheet['C' + str(row)].value
                HOU_LIST = sheet['D' + str(row)].value
                POSTD_LIST = sheet['E' + str(row)].value
                STAFF_LIST = sheet['F' + str(row)].value
                
                TRIM_DEPT_ST = str(DEPT_ID)
                TRIM_DEPT = TRIM_DEPT_ST[0:4]

                file.write("--" + DEPT_NAME +  "\n" + "SELECT Distinct EMPLID," + "\n" + "SAL_ADMIN_PLAN, " + "\n" + "CASE" + "\n" + "WHEN SAL_ADMIN_PLAN IN ('FA09','FA10','FA12','FACM','FA9M','FASU') THEN " + "'" + FAC_LIST + "'" + "\n" + "WHEN SAL_ADMIN_PLAN IN ('FAPD','CPFI') THEN " + "'" + POSTD_LIST + "'" + "\n" + "WHEN SAL_ADMIN_PLAN = 'HOUS' THEN " + "'" + HOU_LIST + "'" + "\n" + "WHEN SAL_ADMIN_PLAN IN ('TA09','TA10','TA12','TAR2','TU1N','TU2E','TU2N','US2E','US2N') THEN " + "'" + STAFF_LIST + "'" + " ELSE 'NEEDS CODE' " + "\n" + "END AS LISTSERV" + "\n" + "FROM PS_JOB A " + "\n" + "WHERE A.EFFDT = (SELECT MAX(A_ED.EFFDT) FROM PS_JOB A_ED WHERE A.EMPLID = A_ED.EMPLID " + "\n" + " AND A.EMPL_RCD = A_ED.EMPL_RCD" + "\n" + "AND A_ED.EFFDT <= SYSDATE)" + "\n" + "AND A.EFFSEQ = (SELECT MAX(A_ES.EFFSEQ) FROM PS_JOB A_ES WHERE A.EMPLID = A_ES.EMPLID" + "\n" + " AND A.EMPL_RCD = A_ES.EMPL_RCD" + "\n" + " AND A.EFFDT = A_ES.EFFDT)" + "\n" + "AND A.EMPL_STATUS IN ('A','L','P','S','W')" + "\n" + "AND A.SAL_ADMIN_PLAN IN ('FA09','FA10','FA12','FACM','FA9M','FASU','FAPD','CPFI','HOUS','TA09','TA10','TA12','TAR2','TU1N','TU2E','TU2N','US2E','US2N')" + "\n" + "AND A.DEPTID LIKE " + "'" + TRIM_DEPT + "%" + "'" + "\n" + "\n" + "UNION" + "\n" + "\n")

def main():
    os.chdir('H:\SavedQueries')
    LIST = "ListServedBuiltQueries.txt"

    build_queries(LIST)


if __name__ == "__main__":
    main()