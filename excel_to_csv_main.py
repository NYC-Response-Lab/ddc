
from azure.storage.blob import ContainerClient
import argparse
import os
import pandas as pd
import io
import csv
import excel_to_csv_convertor as convertor


parser = argparse.ArgumentParser()
parser.add_argument('--url', type=str)
subparsers = parser.add_subparsers()


# Command: test
def test(args):
    print('running the command with for url: %s' % args.url)
    SHEET = 'ddc-data/SampleProject2_02-17-2017_BidVarianceAnalysisDDC.xlsx'
    data = pd.read_excel(SHEET, sheet_name=0, skiprows=7,
                         converters={'RSMeans 12-digit code': lambda x: str(x)}
                         ).fillna('')
    csv_rows = convertor.process_excel_file_as_pd(data, 'PROJECT_ID')

    try:
        with open('test.csv', 'w') as csvfile:
            row_writer = csv.writer(csvfile)
            for row in csv_rows:
                row_writer.writerow(row)
    except Exception:
        print('Error writing csv')


test_cmd = subparsers.add_parser('test')
test_cmd.set_defaults(func=test)


# Command: list_files
def list_files(args):
    container = ContainerClient.from_container_url(args.url)
    for blob in container.list_blobs():
        print(blob['name'])


list_files_cmd = subparsers.add_parser('list_files')
list_files_cmd.set_defaults(func=list_files)


# Command: download files
def download_locally(args):
    assert os.path.isdir(
        args.folder), "folder %s does not exist. You must create it before running the command." % args.folder
    container = ContainerClient.from_container_url(args.url)
    for blob in container.list_blobs():
        filename = os.path.join(args.folder, blob['name'])
        with open(filename, "wb") as my_blob:
            stream = container.download_blob(blob)
            data = stream.readall()
            my_blob.write(data)
            # TODO: add logging


download_locally_cmd = subparsers.add_parser('download_locally')
download_locally_cmd.add_argument('--folder', default=".")
download_locally_cmd.set_defaults(func=download_locally)


# Command: convert to csv
def convert_all_files(args):
    assert os.path.isdir(
        args.folder), "folder %s does not exist. You must create it before running the command." % args.folder
    container = ContainerClient.from_container_url(args.url)
    for blob in container.list_blobs():
        print('Processing %s ...' % blob['name'])
        stream = container.download_blob(blob)
        excel_file = io.BytesIO(stream.readall())
        data = pd.read_excel(excel_file, sheet_name=0,
                             header=None, nrows=1).fillna('')
        project_id = data[4][0]
        if project_id == '':
            print('ERROR processing file %s.' % blob['name'])
            continue
        data = pd.read_excel(excel_file, sheet_name=0, skiprows=7,
                             converters={
                                 'RSMeans 12-digit code': lambda x: str(x)}
                             ).fillna('')

        csv_rows = convertor.process_excel_file_as_pd(data, project_id)

        filename = os.path.join(args.folder, ".".join(
            blob['name'].split('.')[:-1] + ['csv']))

        try:
            with open(filename, 'w') as csvfile:
                row_writer = csv.writer(csvfile)
                for row in csv_rows:
                    row_writer.writerow(row)
        except Exception:
            print('Error writing file to csv')


convert_all_files_cmd = subparsers.add_parser('convert_all')
convert_all_files_cmd.add_argument('--folder', default=".")
convert_all_files_cmd.set_defaults(func=convert_all_files)


# Command: check_unique
def check_unique(args):
    print('checking that each doc has a unique id')
    container = ContainerClient.from_container_url(args.url)
    for blob in container.list_blobs():
        stream = container.download_blob(blob)
        excel_file = io.BytesIO(stream.readall())
        data = pd.read_excel(excel_file, sheet_name=0, header=None, nrows=1)
        project_id = data[4][0]
        print(project_id)
        # TODO: check that all project ids are unique.
        # TODO: add logging


check_unique_cmd = subparsers.add_parser('check_unique')
check_unique_cmd.set_defaults(func=check_unique)


if __name__ == '__main__':
    args = parser.parse_args()
    args.func(args)
