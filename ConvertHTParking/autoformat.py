import re


def format_vehicle_plate_motor(input_str: str) -> str:
    output_lst = []
    for chr in input_str:
        if chr.isalnum():
            output_lst.append(chr.upper())
    if len(output_lst) == 8:
        return "%s%s-%s%s-%s%s%s%s" % tuple(output_lst)
    elif len(output_lst) == 9:
        return "%s%s-%s%s-%s%s%s.%s%s" % tuple(output_lst)
    elif len(output_lst) == 10:
        return "%s%s%s-%s%s-%s%s%s%s%s" % tuple(output_lst)
    return input_str


def format_vehicle_plate_car(input_str: str) -> str:
    output_lst = []
    for chr in input_str:
        if chr.isalnum():
            output_lst.append(chr.upper())
    if len(output_lst) == 8:
        return "%s%s%s-%s%s%s.%s%s" % tuple(output_lst)
    return input_str


def format_card_code(input_str: str) -> str:
    fix = input_str.replace(" ", "").upper()
    out = re.sub(r'(NO|N0)(\.|,)*(?=[0-9A-Z])', "No.", fix)
    return out
