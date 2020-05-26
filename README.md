# Work for NYC DDC

* excel_to_csv_main.py (main program)
* excel_to_csv_convertor.py (code for the conversion)

The main program can:
* download all the files locally (you can specify the destination)
* check that each file has a unique id
* convert all the files into CSV (you can specify the destination)

To convert all the file:
`python excel_to_csv_main.py --url "https://ddcaistorage.blob.core.windows.net/cesamplefiles?st=2020-05-22T16%3A59%3A25Z&se=2020-06-12T16%3A59%3A00Z&sp=racwdl&sv=2018-03-28&sr=c&sig=G7bJsBqqfeduGONJdpXfXeA%2BYa3lTtWQMPpkMoPmeE4%3D"  test`