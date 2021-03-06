from resources.input_file_class import InputFile
from resources.output_file_class import OutputFile

from os import walk, chdir
'''------------------------------------------------------------------------------------------------------------------'''

def main_body(dir = ''):
    print("start")
    print()
    output = OutputFile()

    chdir(dir + '../input')
    for root, dirs, files in walk(".", topdown=False):
        files = files
    for file in files:
        input = InputFile(file, )
        input_dict = input.sheet_values
        output.handle_input_file(input_dict)
        print(output.main_table_rows_values[output.row_id - 2])

    input = None
    output.save_file()
    print()


if __name__ == '__main__':
    main_body()