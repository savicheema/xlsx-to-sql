import os, xlrd, csv
from sqlalchemy import create_engine, Column, String, Integer, MetaData, Table
from sqlalchemy.orm import mapper, create_session

def xl_to_csv(file_name):
    input_dir = os.path.abspath('input')
    output_dir = os.path.abspath('output')
    xlsx_path = os.path.join(input_dir, "%s.xlsx" % file_name)
    csv_path = os.path.join(output_dir,  "%s.csv" % file_name)
    try:
        wb = xlrd.open_workbook(xlsx_path)
    except IOError:
        print(xlsx_path, os.path.exists(xlsx_path))
    sh = wb.sheet_by_index(0)
    your_csv_file = open(csv_path, 'wb')
    # import pdb
    # pdb.set_trace()
    # except Exception:
    #     print(Exception)
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in xrange(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()
    # return you_csv_file


def csv_to_table(file_name, metadata):
    dir_path = os.path.abspath('output')
    # file_path = os.path.abspath("%s.csv" % file_name)
    file_path = os.path.join(dir_path, "%s.csv" % file_name)
    table = None
    # try:
    with open(file_path) as f:
        # assume first line is header
        cf = csv.DictReader(f, delimiter=',')
        for row in cf:
            if table is None:
                # create the table
                table = Table('foo4', metadata, 
                    Column('id', Integer, primary_key=True),
                    *(Column(rowname, String()) for rowname in row.keys()))
                table.create()
            # insert data into the table
            table.insert().values(**row).execute()

    # from subprocess import call
    # call(["mv", file_path, dest_path])
    class CsvTable(object): pass
    mapper(CsvTable, table)


def main():
    import sys
    file_name = sys.argv[1]
    xl_to_csv(file_name)
    # engine = create_engine('sqlite://')
    engine = create_engine("postgresql://savitoj:savi@localhost/savitoj", echo=True)
    metadata = MetaData(bind=engine)
    csv_to_table(file_name, metadata)
    session = create_session(bind=engine, autocommit=False, autoflush=True)


main()
