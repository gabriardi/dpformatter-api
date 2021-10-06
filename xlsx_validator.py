from re import S


def is_xlsx_base64(xlsx_base64):
    if (
        "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"
        not in xlsx_base64
    ):
        return False

    return True


def is_worksheet_valid(ws):
    return True
