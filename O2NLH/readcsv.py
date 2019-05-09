# import configparser
import re
import xlsxwriter
import tempfile
import os


def converter(filepath, journal):

    tmpdir = tempfile.gettempdir()
    basename = os.path.basename(filepath)
    xlname = basename.split(".")[0] + ".xlsx"
    xlfullpath = os.path.join(tmpdir, xlname)
    wb = xlsxwriter.Workbook(xlfullpath)
    ws = wb.add_worksheet("IMPORT")

    col_date    = 0
    col_compte  = 1
    col_tiers   = 2
    col_libelle = 4
    col_debit   = 6
    col_credit  = 7

    with open(filepath, "r") as f:
        csv = f.readlines()

    data = []
    for line in csv:
        # si commence pas par jj/mm/aaa;4|7 on supprime
        if re.search("[0-3][0-9]/[0-1][0-9]/20[0-9][0-9];[47]", line):
            # Damage control pour éviter les IndexError sur les appel de listes
            # si nb colonne inférieur à 8, on supprime
            if len(line.split()) > 8:
                data.append(line)

    rows = []

    for line in data:
        date = ""
        debit = ""
        credit = ""
        lib_1 = ""
        columns = line.split(";")

        # numero de piece
        if " FCPONT" in columns[col_libelle]:
            piece = columns[col_libelle].split(" FCPONT")[1][0:6]
        else:
            piece = ""

        date = columns[col_date]
        compte = columns[col_compte]
        if columns[col_tiers]:
            tiers = columns[col_tiers]
        split_lib = columns[col_libelle].split(" Client ")
        if len(split_lib) == 2:
            lib_1 = split_lib[0]
            lib_2 = split_lib[1]
        debit = float(columns[col_debit].replace(",", "."))
        credit = float(columns[col_credit].replace(",", "."))

        if compte == "411000":
            # CARSAT
            if tiers == "272_CARSATNORMANDIE-760":
                compte = "9272CAR"
                ana = "027"
            # DEPT DE L'EURE
            elif tiers == "272_FRANCHISE-DEPARTEME":
                compte = "9272FRAN"
                ana = "027"
            # O2 FRANCE
            elif tiers == "O2 FRANCE":
                compte = "9O2"
                ana = "027"
            # ANCIEN NUM CLIENT
            elif tiers.startswith("27200"):
                compte = "90" + lib_2[2:]
                if lib_2[2] == "2":
                    ana = "027"
                else:
                    ana = "014"
            # CAS GENERAL
            else:
                compte = "9" + lib_2[6:]
                if lib_2[6:8] == "02":
                    ana = "027"
                else:
                    ana = "014"
            rows.append([journal, date, compte, piece, lib_1, debit, credit, ""])

        elif compte.startswith("445"):
            corresp = {
                "445710" : "445712",
                "445711" : "445713",
                "445713" : "445711",
            }
            rows.append([journal, date, corresp[compte], piece, lib_1, debit, credit, ""])

        elif compte.startswith("7"):
            rows.append([journal, date, compte, piece, lib_1, debit, credit, ana])

        else:
            rows.append([journal, date, compte, piece, lib_1, debit, credit, ""])


    bold = wb.add_format({'top': 1})

    temp_pce = ""
    boldy = False

    for row_index, row in enumerate(rows):
        if row[3] == temp_pce:
            boldy = False
        else:
            boldy = True
        for col_index, col in enumerate(row):
            if boldy:
                ws.write(row_index+1, col_index, col, bold)
            else:
                ws.write(row_index+1, col_index, col)

        temp_pce = row[3]

    wb.close()
    os.system("start excel.exe {}".format(xlfullpath))
    return len(rows)

if __name__ == '__main__':
    # import pprint

    # pp = pprint.PrettyPrinter(indent=4)    
    converter("./assets/ventes.csv", "VE")