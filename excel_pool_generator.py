#!/usr/bin/env python3


import xlsxwriter
import sys


def generate(fp, names):
    print(f"Generating pool in {fp} ...")

    workbook = xlsxwriter.Workbook(fp)
    worksheet = workbook.add_worksheet()

    unused_cell_format = workbook.add_format({'bg_color': 'black'})
    name_cell_format = workbook.add_format({'bold': True})
    n = 1
    for name in names:
        worksheet.write(n, 0, name, name_cell_format)
        worksheet.write(0, n, name, name_cell_format)
        worksheet.write(n, n, "", unused_cell_format)
        n += 1

    worksheet.write(0, n, "Victories",
                    workbook.add_format({'bg_color': 'green'}))
    worksheet.write(0, n+1, "Defeats",
                    workbook.add_format({'bg_color': 'orange'}))

    workbook.close()
    print("Done !")


def get_all_names():
    names = []
    print("\nEnter all participants one by one (press C-D to stop)")

    while True:
        try:
            tmp = input(">> ")
            names.append(tmp)
        except EOFError:
            print(names)
            print("\nAre all participants here (y/n) ?", end=" : ")
            ans = input("")
            if ans != 'y' and ans != "yes":
                print("Continue then\n")
                continue
            else:
                return names
        except :
            print("Bye", file=sys.stderr)
            sys.exit(1)


def main(ac, av):
    if ac == 1:
        print(f"USAGE: {sys.argv[0]} file_dest", file=sys.stderr)
        sys.exit(1)
    file_dest = av[1]

    if not file_dest.endswith(".xlsx"):
        file_dest += ".xlsx"
    print(f"Target file : {file_dest} ok ? (y/n)", end=" : ")
    ans = input()
    if ans != 'y' and ans != "yes":
        print("Bye", file=sys.stderr)
        sys.exit(1)
    names = get_all_names()
    generate(file_dest, names)


if __name__ == '__main__':
    main(len(sys.argv), sys.argv)
