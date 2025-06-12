import pandas
import os
import win32com.client
# import openpyxl
# from classes.control import ctrlString

class ctrlExcel():
    def testexcel():
        result = 0
        
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            version = float(excel.Version)
            excel.Quit()

            # Excel 2019 より前であればエラー
            if version < 16.0:
                result = -5001
        except Exception as err:
            result = -5002

        return result

    def readxlsx(input_path):
        """
        XLSXファイル読み込み
        """
        # ファイル存在チェック
        if not os.path.isfile(input_path):
            return None
        # XLSX読み込み
        try:
            xlsx_data = None
            xlsx_data = pandas.read_excel(input_path,
                                          header=None,
                                          sheet_name=0,
                                          skiprows=1)
        except OSError:
            print('例外エラー：XLSX読み込み')
        return xlsx_data

#     def shapingdata(data_tx, data_ta):
#         """
#         SENDIG v3.1/v.3.1.1のデータ整形
#         """
#         book = openpyxl.Workbook()
#         sheet = book['Sheet']
#         sheet.title = FROM_MAIN
#         # データ読み込み
#         #   TXドメイン
#         flag_add = 0
#         list_tx = []
#         for index, row in data_tx.iterrows():
#             if row[5] == 'SPGRPCD':
#                 flag_add |= FLAG_TX_SPGRPCD
#                 spgrpcd = row[7]
#             elif row[5] == 'ARMCD':
#                 flag_add |= FLAG_TX_ARMCD
#                 armcd = row[7]
#             elif row[5] == 'TKDESC':
#                 flag_add |= FLAG_TX_TKDESC
#                 tkdesc = row[7]
#             if flag_add == FLAG_TX_ADD:
#                 list_tx.append(structDomain.struct_tx(row[0], row[1], row[2],
#                                                       row[3], spgrpcd, armcd,
#                                                       tkdesc))
#                 flag_add = 0
#         # TAドメイン SENDIG v3.1, v3.1.1の場合
#         list_ta = []
#         for index, row in data_ta.iterrows():
#             list_ta.append(structDomain.struct_ta(row[0], row[1], row[2],
#                                                   row[3], row[4], row[5],
#                                                   row[6], row[7], row[8],
#                                                   row[9]))
#         # 紐づけ
#         list_trialdesign = []
#         for row in list_tx:
#             list_trialdesign.append(
#                 structTrialdesign.struct_main(
#                     row.spgrpcd,
#                     row.armcd,
#                     'arm',
#                     'acclimation',
#                     'treatment',
#                     'recovery',
#                     row.setcd,
#                     row.set))
#         row_count = 1
#         # Excelデータへ変換
#         for row in list_trialdesign:
#             sheet.cell(row=row_count, column=1).value = row.spgrpcd
#             sheet.cell(row=row_count, column=2).value = row.armcd
#             sheet.cell(row=row_count, column=3).value = row.arm
#             sheet.cell(row=row_count, column=4).value = row.acclimation
#             sheet.cell(row=row_count, column=5).value = row.treatment
#             sheet.cell(row=row_count, column=6).value = row.recovery
#             sheet.cell(row=row_count, column=7).value = row.setcd
#             sheet.cell(row=row_count, column=8).value = row.set
#             row_count += 1

#         return book

#     def shapingdata_dart(data_tx, data_ta, data_te, data_tp):
#         """
#         SENDIG Dartのデータ整形
#         """
#         book = openpyxl.Workbook()
#         sheet = book['Sheet']
#         sheet.title = FROM_MAIN
#         # データ読み込み
#         #   TXドメイン
#         flag_add = 0
#         list_tx = []
#         for index, row in data_tx.iterrows():
#             if row[5] == 'SPGRPCD':
#                 flag_add |= FLAG_TX_SPGRPCD
#                 spgrpcd = row[7]
#             elif row[5] == 'ARMCD':
#                 flag_add |= FLAG_TX_ARMCD
#                 armcd = row[7]
#             elif row[5] == 'TKDESC':
#                 flag_add |= FLAG_TX_TKDESC
#                 tkdesc = row[7]
#             if flag_add == FLAG_TX_ADD:
#                 list_tx.append(structDomain.struct_tx(row[0], row[1], row[2],
#                                                       row[3], spgrpcd, armcd,
#                                                       tkdesc))
#                 flag_add = 0
#         #   TAドメイン SENDIG Dartの場合
#         list_ta = []
#         for index, row in data_ta.iterrows():
#             list_ta.append(structDomain.struct_ta_dart(row[0], row[1],
#                                                        row[2], row[3],
#                                                        row[4], row[5],
#                                                        row[6], row[7]))
#         #   TEドメイン
#         list_te = []
#         for index, row in data_te.iterrows():
#             list_te.append(structDomain.struct_te(row[0], row[1],
#                                                   row[2], row[3],
#                                                   row[4], row[5],
#                                                   row[6]))
#         #   TPドメイン
#         row_count = 0
#         list_tp = list(range((len(data_tp))))
#         for index, row in data_tp.iterrows():
#             list_tp[row_count] = structDomain.struct_tp(row[0], row[1],
#                                                         row[2], row[3],
#                                                         row[4], row[5],
#                                                         row[6], row[7],
#                                                         row[8])
#         # 紐づけ
#         list_trialdesign = []
#         for row in list_tx:
#             list_trialdesign.append(
#                 structTrialdesign.struct_main(
#                     row.spgrpcd,
#                     row.armcd,
#                     'arm',
#                     'acclimation',
#                     'treatment',
#                     'recovery',
#                     row.setcd,
#                     row.set))
#         row_count = 1
#         # Excelデータへ変換
#         for row in list_trialdesign:
#             sheet.cell(row=row_count, column=1).value = row.spgrpcd
#             sheet.cell(row=row_count, column=2).value = row.armcd
#             sheet.cell(row=row_count, column=3).value = row.arm
#             sheet.cell(row=row_count, column=4).value = row.acclimation
#             sheet.cell(row=row_count, column=5).value = row.treatment
#             sheet.cell(row=row_count, column=6).value = row.recovery
#             sheet.cell(row=row_count, column=7).value = row.setcd
#             sheet.cell(row=row_count, column=8).value = row.set
#             row_count += 1

#         return book

#     def writeexcel(book):
#         """
#         XLSXファイルの書き込み
#         """
#         result = 0
#         # データ存在チェック
#         if book is None:
#             result = -301
#             return result
#         downloaddir = os.path.expanduser(r'~\Downloads')
#         datestr = ctrlString.now_filename()
#         filename_array = [downloaddir,
#                           '\\',
#                           'TrialDesignForm_',
#                           datestr,
#                           '.xlsx']
#         filename_output = ''.join(filename_array)

#         try:
#             book.save(filename=filename_output)
#             book.close()
#         except OSError:
#             result = -311
#             print('例外エラー：Excel書き込み')
#         return result

#     def gettrialtype(sendversion):
#         """
#         SENDバージョン判定しラジオボタンで使用する値を返す
#         """
#         if sendversion == TSVAL_V3_1:
#             trialtype = 1
#         elif sendversion == TSVAL_V3_1_1:
#             trialtype = 1
#         elif sendversion == TSVAL_DART:
#             trialtype = 2
#         else:
#             trialtype = 0
#         return trialtype
