import os
from PyPDF2 import PdfFileReader, PdfFileWriter


class PdfHelper:

    def __init__(self, file_name):
        self.file_name = file_name
        self.reader = PdfFileReader(self.file_name, strict=False)
        self.num_pages = self.reader.getNumPages()
        if self.reader.isEncrypted:
            print('Trying to decrypt ...')
            try:
                self.reader.decrypt('')
                print('Success!')
            except:
                print('Failed to decrypt')

    def split_pages(self, start_page, end_page, output_name):
        if start_page <= end_page <= self.num_pages:
            self.select_pages(range(start_page, end_page + 1), output_name)
        else:
            print('ERROR: page number out of range: start {}, end {}, total {}.'.format(start_page, end_page,
                                                                                        self.num_pages))

    def select_pages(self, pages, output_name):
        if max(pages) - 1 <= self.num_pages:
            writer = PdfFileWriter()
            for p in pages:
                writer.addPage(self.reader.getPage(p - 1))
            with open(output_name, 'wb') as outfile:
                writer.write(outfile)
        else:
            print('ERROR: page number out of range: max {}, total {}.'.format(max(pages), self.num_pages))


if __name__ == '__main__':
    file_base_name = input('文件前缀（年份之前的）：')
    start_year = int(input('起始年份：'))
    end_year = int(input('结束年份：'))
    output_names = []
    # for i in range(3):
    #     output_names.append(input('输出文件名（第{}部分，不包括后缀）：'.format(i + 1)))
    output_names.append("1. Business")
    output_names.append("2. Management's discussion and analysis of financial condition and results of operations")
    output_names.append("3. Financial statements and supplementary data")
    for year in range(start_year, end_year + 1):
        try:
            print('=====================================================')
            print('{}年'.format(year))
            input_file_name = file_base_name + str(year) + '.pdf'
            ph = PdfHelper(input_file_name)
            folder_path = './{}'.format(year)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            for i in range(3):
                print('-----------------------------------------------------')
                print('\t第{}部分'.format(i + 1))
                start_page = int(input('\t\t起始页：'))
                end_page_with_suffix = input('\t\t结束页：（可追加： 25+34 或 25+34~38）\n\t\t        ')
                output_file_name = folder_path + '/' + output_names[i] + '.pdf'
                append_pages = []
                if '+' in end_page_with_suffix:
                    [end_page, suffix] = end_page_with_suffix.split('+')
                    if suffix and '~' in suffix:
                        pages = map(int, suffix.split('~'))
                        append_pages += pages
                    else:
                        append_pages += [int(suffix)]
                    end_page = int(end_page)
                else:
                    end_page = int(end_page_with_suffix)
                page_range = list(range(start_page, end_page + 1)) + append_pages
                ph.select_pages(page_range, output_file_name)
                # ph.split_pages(start_page - 1, end_page - 1, output_file_name)
            print('=====================================================')
        except Exception as e:
            print('失败：' + file_base_name + str(year) + '.pdf')
            print(e)
