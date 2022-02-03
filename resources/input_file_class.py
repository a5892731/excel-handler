from resources.file_class import File

class InputFile(File):
    def __init__(self, filename, path = '../input', sheet_name='Arkusz1'):
        '''this is input file class'''
        super().__init__(filename, path, sheet_name)

        '''sheet variables'''
        self.sheet_coordinates = {'Klient': 'H1',
                                  'Adres':'H2',
                                  'Faktura': 'H3',
                                  'Przesyłka': 'H4',
                                  'Status': 'H5',
                                  'Transakcja': 'H6',
                                  'Suma': 'H7',
                                  }
        self.sheet_values = {}
        self.get_data()

    def get_data(self):
        def get_row_and_column(coordinates):
            '''function get column and row from excel coordinates like in example: AA77 -> column: AA, row: 77'''
            column = ''
            row = ''
            for sign in coordinates:
                if sign.isalpha():
                    column += sign
                if sign.isnumeric():
                    row += sign
            return column, int(row)

        def get_sum(coordinates, sheet_ranges, splitsign = ':'):
            '''this function gets sum of elements from row coordinates like in example: A1:A22 -->> =SUM(A1:A22)'''
            start, stop = coordinates[5:-1].split(splitsign)
            column, start_row = get_row_and_column(start)
            column, stop_row = get_row_and_column(stop)
            sum = 0
            for row in range(start_row, stop_row + 1):
                cell_data = sheet_ranges[column + str(row)].value
                value, opertion = get_mathematical_operation(cell_data, sheet_ranges)
                if value == None:
                    value = 0
                sum += value
            return sum

        def get_mul(coordinates, sheet_ranges, splitsign = '*'):
            '''this function gets sum of elements from row coordinates like in example: =C2*D2 -->> C2 and D2'''
            start, stop = coordinates[1:].split(splitsign)
            column1, row1 = get_row_and_column(start)
            column2, row2 = get_row_and_column(stop)
            a = sheet_ranges[column1 + str(row1)].value
            a, opertion = get_mathematical_operation(a, sheet_ranges)
            b = sheet_ranges[column2 + str(row2)].value
            b, opertion = get_mathematical_operation(b, sheet_ranges)
            return int(a) * int(b)

        def get_mathematical_operation(cell_data, sheet_ranges):
            '''in future use ->>>>>>                                from openpyxl.formula.translate import Translator'''
            if '=SUM' in str(cell_data):
                sum = get_sum(cell_data, sheet_ranges)
                return sum, 'sum'
            elif '=' in str(cell_data) and '*' in str(cell_data):
                mul = get_mul(cell_data, sheet_ranges)
                return mul, 'mul'
            else:
                return cell_data, 'assignment' # in

        '''----------------------------------------------------------------------------------------------------------'''
        sheet_ranges = self.wb[self.sheet_name]

        for key in self.sheet_coordinates:
            value = sheet_ranges[self.sheet_coordinates[key]].value

            value, opertion = get_mathematical_operation(value, sheet_ranges)
            self.sheet_values[key] = value

'''------------------------------------------------------------------------------------------------------------------'''
if __name__ == '__main__':
    print("start")
    print()

    file = InputFile('file 1.xlsx')
    print(file.sheet_values)

    print()