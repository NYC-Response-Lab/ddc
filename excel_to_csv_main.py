
from azure.storage.blob import ContainerClient
import argparse
import os
import pandas as pd
import io
import csv
import excel_to_csv_convertor as convertor
import logging
import sys

logger = logging.getLogger(__name__)
handler = logging.StreamHandler(sys.stderr)
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.setLevel(logging.INFO)
logger.addHandler(handler)


parser = argparse.ArgumentParser()
parser.add_argument('--url', type=str)
subparsers = parser.add_subparsers()

# Command: test


def test(args):
    print('running the command with for url: %s' % args.url)
    SHEET = 'ddc-data/SampleProject2_02-17-2017_BidVarianceAnalysisDDC.xlsx'
    # SHEET = 'SampleProject7_07-12-2019_BidVarianceAnalysisDDC.xlsx'
    SHEET = 'SampleProject8_12-21-2017_BidVarianceAnalysisDDC.xlsx'
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
    logger.info('Listing files from %s.' % args.url)
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
        logger.info('Processing file %s.' % blob['name'])
        filename = os.path.join(args.folder, blob['name'])
        with open(filename, "wb") as my_blob:
            stream = container.download_blob(blob)
            data = stream.readall()
            my_blob.write(data)
            logger.info('File saved as %s.' % filename)


download_locally_cmd = subparsers.add_parser('download_locally')
download_locally_cmd.add_argument('--folder', default=".")
download_locally_cmd.set_defaults(func=download_locally)


# Command: convert to csv
def convert_all_files(args):
    assert os.path.isdir(
        args.folder), "folder %s does not exist. You must create it before running the command." % args.folder
    ERRORS = []
    METADATA_ROWS = []
    container = ContainerClient.from_container_url(args.url)
    for blob in container.list_blobs():
        try:
            logger.info('Processing %s ...' % blob['name'])
            stream = container.download_blob(blob)
            excel_file = io.BytesIO(stream.readall())

            # We load the sheet a first time to extract metaadata.
            # We only need the first 4 rows (`nrows=4`).
            data = pd.read_excel(excel_file, sheet_name=0,
                                 header=None, nrows=4).fillna('')
            project_id = data[4][0]
            if project_id == '':
                logger.error('Cannot find project_id for `%s`.' %
                             blob['name'])
                ERRORS.append(blob['name'])
                continue
            bid_date = data[4][1]
            bid_comparison_date = data[4][2]
            project_name = data[4][3]
            ddc_engineer_estimate = None
            first_bidder = None
            second_bidder = None
            third_bidder = None

            METADATA_ROWS.append([project_id, project_name, bid_date,
                                  bid_comparison_date, ddc_engineer_estimate, first_bidder, second_bidder, third_bidder])

            # We load the sheet a second time to parse the main table.
            data = pd.read_excel(excel_file, sheet_name=0, skiprows=7,
                                 converters={
                                     'RSMeans 12-digit code': lambda x: str(x)}
                                 ).fillna('')

            csv_rows = convertor.process_excel_file_as_pd(data, project_id)

            filename = os.path.join(args.folder, ".".join(
                blob['name'].split('.')[:-1] + ['csv']))

            logger.info('Writing to csv.')
            try:
                with open(filename, 'w') as csvfile:
                    row_writer = csv.writer(csvfile)
                    for row in csv_rows:
                        row_writer.writerow(row)
            except Exception:
                logger.error('Error writing file `%s` to csv' % blob['name'])
                ERRORS.append(blob['name'])
        except Exception as e:
            logger.error('Problem with file %s.' % blob['name'])
            logger.error(e)
            print(e)
            ERRORS.append(blob['name'])

    filename = os.path.join(args.folder, "all_projects.csv")
    with open(filename, 'w') as csvfile:
        row_writer = csv.writer(csvfile)
        for row in METADATA_ROWS:
            row_writer.writerow(row)
    print('Files that could not be processed: %s.' % ",".join(ERRORS))


convert_all_files_cmd = subparsers.add_parser('convert_all')
convert_all_files_cmd.add_argument('--folder', default=".")
convert_all_files_cmd.set_defaults(func=convert_all_files)


# Command: check_unique
def check_unique(args):
    print('checking that each doc has a unique id')
    container = ContainerClient.from_container_url(args.url)
    for blob in container.list_blobs():
        logger.info('Processing file `%s`.' % blob['name'])
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
