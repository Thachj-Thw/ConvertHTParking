from ConvertHTParking import ZKTeck
from ConvertHTParking.autoformat import format_vehicle_plate_car


convert_vehicle = {
    "Thẻ cư dân":               "THẺ CƯ DÂN",
    "Thẻ tháng ô tô":           "XE MÁY THÁNG",
    "Thẻ tháng xe đạp điện":    "XĐ ĐIỆN THÁNG",
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

user_vehicle = [
    "THẺ CƯ DÂN",
    "XE MÁY THÁNG",
    "XĐ ĐIỆN THÁNG",
    "XE ĐẠP THÁNG",
    "XE MÁY VIP",
    "ÔTÔ THÁNG",
    "ÔTÔ VIP"
]

moto = ZKTeck("Input/Danh-sách-thẻ-140820241417-XM, XĐ,... BQL PHỤ TRÁCH.xlsx", convert_vehicle, user_vehicle)
moto.convert("Output/xemay/Dang ki the xm.xlsx", "Output/xemay/Dang ki khach hang xm.xlsx")

convert_vehicle.pop("Thẻ cư dân")
car = ZKTeck("Input/Danh-sách-thẻ-140820241421- ÔTÔ - CĐT PHỤ TRÁCH.xlsx", convert_vehicle, user_vehicle, format_vehicle_plate_method=format_vehicle_plate_car)
car.convert("Output/oto/Dang ki the oto.xlsx", "Output/oto/Dang ki khach hang oto.xlsx")
