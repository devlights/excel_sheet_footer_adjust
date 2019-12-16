#################################################################
# 指定されたフォルダ配下のExcelを開いていき特定の条件にマッチするシートのフッターを調整します.
#
# 実行には、以下のライブラリが必要です.
#   - win32com
#     - $ python -m pip install pywin32
#
# [参考にした情報]
#   - http://excel.style-mods.net/tips_vba/tips_vba_7_09.htm
#   - https://stackoverflow.com/a/37635373
#   - https://www.sejuku.net/blog/23647
#################################################################
import argparse


# noinspection SpellCheckingInspection
def go(target_dir: str, pattern: str, footer: str):
    import pathlib

    import pywintypes
    import win32com.client

    excel_dir = pathlib.Path(target_dir)
    if not excel_dir.exists():
        print(f'target directory not found [{target_dir}]')
        return

    try:
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True

        for f in excel_dir.glob('**/*.xlsx'):
            abs_path = str(f)
            try:
                wb = excel.Workbooks.Open(abs_path)
            except pywintypes.com_error as err:
                print(err)
                continue

            try:
                sheets_count = wb.Sheets.Count
                for sheet_index in range(0, sheets_count):
                    ws = wb.Worksheets(sheet_index + 1)
                    ws.Activate()
                    if not pattern:
                        ws.PageSetup.CenterFooter = footer
                    else:
                        if pattern in ws.Name:
                            ws.PageSetup.CenterFooter = footer
                if sheets_count >= 0:
                    ws = wb.Worksheets(1)
                    ws.Activate()
                wb.Save()
                wb.Saved = True
            finally:
                wb.Close()
    finally:
        excel.Quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        usage='python main.py -d /path/to/excel/dir -p シート名条件 -f &P',
        description='Excelの特定シートのフッターを調整します.',
        add_help=True
    )

    parser.add_argument('-d', '--directory', help='対象ディレクトリ', required=True)
    parser.add_argument('-p', '--pattern', help='シート名の条件 (python の in 演算子で判定しています）指定しない場合は全シートが対象', default='')
    parser.add_argument('-f', '--footer', help='フッターに設定する値', default='&P')

    args = parser.parse_args()

    go(args.directory, args.pattern, args.footer)
