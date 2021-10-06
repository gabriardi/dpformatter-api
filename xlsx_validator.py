from re import S


def is_xlsx_base64(xlsx_base64):
    if (
        "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"
        not in xlsx_base64
    ):
        return False

    return True


def is_worksheet_valid(ws):
    is_valid = True

    identifiers = [
        {"name": "Diensteinteilervorschlag", "number": 0},
        {"name": "TOZ", "number": 0},
        {"name": "Total bis Saisonende", "number": 0},
        {"name": "Kommentar", "number": 0},
    ]
    for row in ws.iter_rows():
        for cell in row:
            for identifier in identifiers:
                if identifier["name"] in str(cell.value):
                    identifier["number"] += 1
    for identifier in identifiers:
        if identifier["number"] == 0:
            is_valid = False

    return is_valid
