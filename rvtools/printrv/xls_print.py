"""Module responsible to save the content as csv file"""

import xlsxwriter


def xls_print(module_name, object_list, directory):
    """
    Def which will receive the server list and then will
    populate all fields from header as the content.

    Note. All fields will be populated automatically, then according to
    the module which will call csv_print, the # of fields will be updated
    automatically.
    """

    full_file_path = directory + "/" + module_name

    print("=====================")
    print("## Creating {} file.".format(full_file_path))
    print("=====================")

    with xlsxwriter.Workbook(full_file_path) as workbook:
        row_offset = 0
        worksheet = workbook.add_worksheet("all")
        for obj in object_list:
            row_offset += 2
            worksheet.write(row_offset, 0, obj["object"])
            row_offset += 1
            # worksheet = workbook.add_worksheet(obj["object"])
            for section in obj["data"]:
                fields_count = 0
                for col, fields in enumerate(section):
                    for row, (key, value) in enumerate(fields.items()):
                        worksheet.write(row_offset + row, 0, key)
                        worksheet.write(row_offset + row, col + 1, str(value))
                    fields_count = max(fields_count, row + 1)
                row_offset += fields_count


# Write a total using a formula.
# worksheet.write(row, 0, 'Total')
# worksheet.write(row, 1, '=SUM(B1:B4)')
