# author: a5892731
# date: 03.02.2022
# last update: 03.02.2022
# version: 1.0.0
#
# description:
# This program copies the data from the selected cell, of many the same excel files, and then builds the output sheet
# containing the column with the previously copied data.
#


from resources.__main__ import main_body
from os import chdir



if __name__ == '__main__':
    chdir('resources')
    main_body()