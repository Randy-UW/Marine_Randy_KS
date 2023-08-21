import auto_processor as auto
import transfer_conversion_list as transfer
import build_transpose
import os
import copy_and_paste as cap


# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    if input('Do you want to build transpose files?') == 'yes':
        anno_file_path = input('What is the name of file folder for anno file?')
        anno_file_path = "C:\\Users\\Randy\\Desktop\\UW\\" \
                         "KS_Auto_Excel\\" + anno_file_path
        for anno_file in os.listdir(anno_file_path):
            build_transpose.build_transpose_file(anno_file_path + '\\'
                                                 + anno_file, 'Sheet2', 0)
    x = input('new or old files needed to be processed to transversion list file?')
    if x == 'new':
        transfer.transfer_conversion('transpose.xlsx', 'Sheet1',
                                     'ConversionListJune2007.xlsx',
                                     anno_sheet='Final',
                                     anno_file='annotationsJune2007.xls',
                                     temp_file='SessileTemp.xlsx',
                                     temp_sheet_column='Sheet7!M',
                                     temp_quad='Sheet7!K2')
    elif x == 'old':
        file_folder = input('Which transpose file folder need to build conversion list file?')
        transpose_path = "C:\\Users\\Randy\\Desktop\\UW\\" \
                         "KS_Auto_Excel\\"
        prefix = input('What is the prefix?')
        suffix = input('What is the suffix?')
        transpose_path += file_folder
        for transpose_file in os.listdir(transpose_path):
            transfer.transfer_old_conversion(transpose_path + '\\' + transpose_file, 'Sheet',
                                             anno_sheet='Sheet2',
                                             anno_file=transpose_file,
                                             temp_file='SessileMAtemplSpt22old.xlsx',
                                             temp_sheet_column='Sheet9!AA',
                                             temp_quad='Sheet9!C2',
                                             prefix=prefix,
                                             suffix=suffix)

    if input('Transfer the conversion list to the file we need?') == 'yes':
        print('put the anno files to the main directory')
        path = input('Conversion List\'s file folder name')
        path = "C:\\Users\\Randy\\Desktop\\UW\\KS_Auto_Excel\\" + path
        for file in os.listdir(path):
            file = path + '\\' + file
            print(file)
            auto.open_excel(file, 'Sheet')

    if input('Copy and paste for some columns?') == 'yes':
        copy_path = input('What is the copy file folder?')
        paste_path = input('What is the paste file folder?')
        copy_path = "C:\\Users\\Randy\\Desktop\\UW\\KS_Auto_Excel\\" + copy_path
        paste_path = "C:\\Users\\Randy\\Desktop\\UW\\KS_Auto_Excel\\" + paste_path
        for copy_file in os.listdir(copy_path):
            for paste_file in os.listdir(paste_path):
                if copy_file[0:9] == paste_file[0:9]:
                    cap.copy_and_paste_col(copy_path+"\\"+copy_file, paste_path+"\\"+paste_file)
                    print(paste_file + ' completed!')
