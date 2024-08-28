from ConvertHTParking import KzParking
from ConvertHTParking.autoformat import format_vehicle_plate_car, format_vehicle_plate_motor


convert_vehicle = {
    "Thẻ cư dân":               "THẺ CƯ DÂN",
    "Thẻ tháng xe máy":         "XE MÁY THÁNG",
    "Thẻ tháng xe đạp điện":    "XĐ ĐIỆN THÁNG",
    "Thẻ tháng xe đạp":         "XE ĐẠP THÁNG",
    "Thẻ VIP xe máy":           "XE MÁY VIP",
    "Thẻ VIP xe đạp điện":      "XĐ ĐIỆN VIP",
    "Thẻ VIP xe đạp":           "XE ĐẠP VIP",
    "Thẻ lượt xe máy":          "XE MÁY LƯỢT",
    "Thẻ lượt xe đạp điện":     "XĐ ĐIỆN LƯỢT",
    "Thẻ lượt xe đạp":          "XE ĐẠP LƯỢT",
    "Thẻ tháng ô tô":           "ÔTÔ THÁNG",
    "Thẻ VIP ô tô":             "ÔTÔ VIP",
    "Thẻ lượt ô tô":            "ÔTÔ LƯỢT"
}

vehicle_month = [
    "THẺ CƯ DÂN",
    "XE MÁY THÁNG",
    "XĐ ĐIỆN THÁNG",
    "XE ĐẠP THÁNG",
    "XĐ ĐIỆN VIP",
    "XE ĐẠP VIP",
    "XE MÁY VIP",
    "ÔTÔ THÁNG",
    "ÔTÔ VIP"
]

prioritize = [
    "ÔTÔ THÁNG", 
    "XE MÁY THÁNG", 
    "XĐ ĐIỆN THÁNG", 
    "XE ĐẠP THÁNG",
    "ÔTÔ VIP", 
    "XE MÁY VIP", 
    "XĐ ĐIỆN VIP",
    "XE ĐẠP VIP",
    "ÔTÔ LƯỢT",
    "XE MÁY LƯỢT",
    "XĐ ĐIỆN LƯỢT",
    "XE ĐẠP LƯỢT",
    "THẺ CƯ DÂN"
]
pass_vehicle = "THẺ CƯ DÂN"
pass_index = 0


def format_plate(plate: str, obj: dict) -> str:
    if "ô tô" in obj["vehicle"]:
        return format_vehicle_plate_car(plate)
    return format_vehicle_plate_motor(plate)


def fix_card_duplicate(card_no: str, obj: dict, invalid: list) -> str:
    if "ÔTÔ" in obj["vehicle"]:
        card_no += "O"
    elif "THÁNG" in obj["vehicle"]:
        card_no += "T"
    if card_no in invalid:
        new_card_no = card_no
        for i in range(65, 90):
            if new_card_no in invalid:
                new_card_no = card_no + chr(i)
            else:
                return new_card_no
    else:
        return card_no
    return fix_card_duplicate(card_no, obj, invalid)



kz = KzParking(convert_vehicle, vehicle_month, prioritize, fix_card_duplicate, format_vehicle_plate_method=format_plate)

kz.convert([
    "Input/Danh-sách-thẻ-140820241421- ÔTÔ - CĐT PHỤ TRÁCH.xlsx",
    "Input/Danh-sách-thẻ-140820241417-XM, XĐ,... BQL PHỤ TRÁCH.xlsx",
    "Input/Mau dang ky khach hang 14.08.xlsx"
    ], 
    "Output/Dang ky the.xlsx", "Output/Dang ky khach hang.xlsx", pass_index, pass_vehicle)