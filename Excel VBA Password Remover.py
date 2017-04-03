from zipfile import ZipFile
from zipfile import is_zipfile
import time
import sys
from os.path import isfile

def main():
    out_filename = "xl/vbaProject.bin"

    try:
        f = sys.argv[1]
    except Exception:
        print('Wrong Arguments')
        return

    if is_zipfile(f):
        try:
            with ZipFile(f, 'r') as zip_in:
                try:
                    with ZipFile("No_PW.xlsm", 'w') as zip_out:
                        try:
                            zip_out.comment = zip_in.comment
                            for in_filename in zip_in.infolist():
                                if out_filename != in_filename.filename:
                                    zip_out.writestr(in_filename, zip_in.read(in_filename))
                                else:
                                    vbafile = zip_in.read(in_filename)
                                    if (vbafile.find(b'\x44\x50\x42') != -1):
                                        vbafile = vbafile.replace(b'\x44\x50\x42', b'\x44\x50\x78')
                                    else:
                                        print('File seems not to be password protected')
                                    zip_out.writestr(in_filename, vbafile)
                                    print('Done!')
                        except Exception:
                            print('Someting Wong')
                except Exception:
                    print('Could not open/ create Output file')
        except Exception:
            print('Could not open input file')
    else:
        if(isfile(f)):
            print('Input file is not a macro enabled workbook')
        else:
            print('Input file does not exist/ is accessible')

if __name__ == '__main__':
    main()
