from resources.file_class import File

class OutputFile(File):
    def __init__(self, filename = 'output_pattern.xlsx', path = '../resources', sheet_name='Arkusz1'):
        '''this is output file class'''
        super().__init__(filename, path, sheet_name)

        '''sheet variables'''
        self.main_table_coordinates = {'Id':'A',
                                       'Transakcja':'B',
                                       'Faktura': 'C',
                                       'Klient': 'D',
                                       'Adres': 'E',
                                       'Przesyłka': 'F',
                                       'Status': 'G',
                                       'Suma': 'H',

                                      }
        self.row_id = 1
        self.main_table_rows_values = []
        self.summary_table_coordinates = {'Ilość Transakcji': 'N1',
                                          'Do zapłaty': 'N2',
                                          'Do zwrotu': 'N3',
                                          'Zysk': 'N4',
                                          }
        self.summary_table_values = {}

    def handle_input_file(self, input_dict):
        '''this function contains all actions for one input file'''

        def get_input_row_data(input_dict):
            input_dict['Id'] = self.row_id
            self.main_table_rows_values.append(input_dict)
            self.row_id += 1

        def add_row_to_work_sheet():
            self.row = 2
            for row_dict in self.main_table_rows_values:
                for key in self.main_table_coordinates:
                    column = self.main_table_coordinates[key]
                    row = str(self.row)
                    self.ws[column + row] = row_dict.get(key)
                self.row += 1
        '''------------------------------------------------------'''
        get_input_row_data(input_dict)
        add_row_to_work_sheet()

    def save_file(self, filename = 'output_file.xlsx', path = '../output', sheet_name='Arkusz1'):
        self.wb.save(filename = path + '/' + filename)

    def __del__(self):
        '''save file'''
        print('output saved')   # - <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< TO DO

'''------------------------------------------------------------------------------------------------------------------'''

if __name__ == '__main__':
    print("start")
    print()

    input_dict = {'Adres': 'Gdańsk ul. Długa 1',
                  'Klient': 'Andrzej Kowalski',
                  'Faktura': 'Wystawiona',
                  'Przesyłka': 'Dostarczona',
                  'Status': 'Opłacone',
                  'Transakcja': 20220203131100,
                  'Suma': 77}

    file = OutputFile()
    file.handle_input_file(input_dict)
    print(file.main_table_rows_values)

    file.save_file()

    print()