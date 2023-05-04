from logging.handlers import RotatingFileHandler
from xlrd import open_workbook
import re
from matplotlib import pyplot as plt
import logging
import argparse


def read_excel(filename):
    """
    reads an excel file and returns the cells of the second column
    :param filename: the name of the file
    :return: the cells of the second column
    """
    try:
        wb = open_workbook(filename)
        sheet = wb.sheet_by_index(0)
        for row in range(1, sheet.nrows):
            yield sheet.cell(row, 1).value
    except FileNotFoundError:
        logging.getLogger('path_creation').warning("File not found")


def split_coord(cell):
    """
    splits the coordinates in the cell
    :param cell: the cell
    :return: the coordinates
    """
    coord = []
    matches = re.finditer("(\d+\.\d+) (\d+\.\d+)", cell)
    for match in matches:
        coord.append([float(match.group(1)), float(match.group(2))])
    return coord


if __name__ == '__main__':
    parser = argparse.ArgumentParser()

    parser.add_argument("infile", help="data_file", type=str)
    parser.add_argument("outfile", help="output_file", type=str)
    parser.add_argument("-v", "--verbosity", help="increase output verbosity", action="store_true")

    args = parser.parse_args()

    logger = logging.getLogger('path_creation')
    logger.setLevel(logging.DEBUG if args.verbosity else logging.WARNING)

    rfh = RotatingFileHandler('logging.log', maxBytes=10000, backupCount=5, encoding="utf-8")
    rfh.setLevel(logger.getEffectiveLevel())

    ch = logging.StreamHandler()
    ch.setLevel(logger.getEffectiveLevel())

    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    rfh.setFormatter(formatter)
    ch.setFormatter(formatter)

    logger.addHandler(rfh)
    logger.addHandler(ch)

    logger.info("Creating plot")
    plt.figure(figsize=(10, 6), dpi=200)
    logger.info("Reading file")

    for cell in read_excel(args.infile):
        path_coords = split_coord(cell)
        plt.plot([x[0] for x in path_coords], [y[1] for y in path_coords])

    plt.title('Hiking Trails')
    plt.savefig(args.outfile)
