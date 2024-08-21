from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.styles import borders
from openpyxl.cell import Cell
import module #module-thw
from .structs import InputForm, OutputUserForm, OutputCardForm, DataStruct
from .autoformat import format_card_code, format_vehicle_plate_motor
from .convert import convert8to10
import re
import os


path = module.Path(__file__)


class ZKTeck:

    input_form = InputForm(
        card_no         = "B",
        card_id         = "C",
        vehicle         = "D",
        vehicle_plate   = "E",
        end_time        = "F",
        status          = "G",
        user_id         = "C",
        name            = "I",
        address         = "J",
        start           = 9
    )

    card_form = OutputCardForm(
        ordinal     = "A",
        card_no     = "B",
        card_id     = "C",
        card_type   = "D",
        vehicle     = "E",
        status      = "G",
        start       = 6
    )

    user_form = OutputUserForm(
        ordinal     = "A",
        name        = "B",
        address     = "C",
        user_id     = "E",
        end_time    = "K",
        vehicle     = "P",
        card_no     = "W",
        start       = 5
    )

    def __init__(self, file_path: str, vehicle_convert: dict, user_vehicle: list, end_row: int = 10000, emty: int = 0, 
                 format_card_method: callable = format_card_code, 
                 format_vehicle_plate_method: callable = format_vehicle_plate_motor, 
                 convert_card_id: callable = convert8to10) -> None:
        self.format_card_method = format_card_method
        self.format_vehicle_plate_method = format_vehicle_plate_method
        self.convert_card_id = convert_card_id
        self.vehicle_convert = vehicle_convert
        self.user_vehicle = user_vehicle
        self.end_row = end_row
        self.emty = emty
        print("loading workbook")
        self.workbook = load_workbook(filename=file_path)
        self.sheet = self.workbook.active
        self.card_file = load_workbook(filename=path.source.join("templates", "BaseCard.xlsx"))
        self.card_file_lock = load_workbook(filename=path.source.join("templates", "BaseCard.xlsx"))
        self.card_user = load_workbook(filename=path.source.join("templates", "BaseUser.xlsx"))
        self.card_user_lock = load_workbook(filename=path.source.join("templates", "BaseUser.xlsx"))
        self.card_sheet = self.card_file.active
        self.user_sheet = self.card_user.active
        self.card_sheet_lock = self.card_file_lock.active
        self.user_sheet_lock = self.card_user_lock.active

    def _split_data(self) -> list[DataStruct]:
        all_data = []
        card_day = []
        emty_counter = 0
        for row in range(self.input_form.start, self.end_row):
            print("Read row: ", row)
            str_row = str(row)
            vehicle = self.sheet[self.input_form.vehicle + str_row].value
            card_no = self.sheet[self.input_form.card_no + str_row].value
            card_id = self.sheet[self.input_form.card_id + str_row].value
            vehicle_plate = self.sheet[self.input_form.vehicle_plate + str_row].value
            end_time = self.sheet[self.input_form.end_time + str_row].value
            status = self.sheet[self.input_form.status + str_row].value
            name = self.sheet[self.input_form.name + str_row].value
            address = self.sheet[self.input_form.address + str_row].value
            if not vehicle and not card_no and not card_id and not vehicle_plate and not end_time and not status and not name and not address:
                emty_counter += 1
                if emty_counter > self.emty:
                    break
            if vehicle in self.vehicle_convert:
                vehicle = self.vehicle_convert[vehicle]
            else:
                continue
            if self.format_card_method:
                card_no = self.format_card_method(card_no)
            if not card_no or not card_id.isnumeric():
                continue
            if self.convert_card_id:
                card_id = self.convert_card_id(card_id)
            if self.format_vehicle_plate_method:
                vehicle_plate = self.format_vehicle_plate_method(vehicle_plate)
            if self.input_form.user_id == self.input_form.card_id:
                user_id = card_id
            else:
                user_id = self.sheet[self.input_form.user_id + str_row].value
            is_month = vehicle in self.user_vehicle
            if not is_month:
                card_day.append(card_no)
            data = DataStruct(card_no, card_id, user_id, True if status == "Hoạt động" else False ,vehicle, vehicle_plate, end_time, name, address, is_month)
            all_data.append(data)
        
        for data in all_data:
            if data.is_month and data.card_no in card_day:
                data.card_no += "T"

        print("sorting")
        all_data.sort(key=lambda x: (len(x.card_no), x.card_no))
        return all_data

    def convert(self, save_card_as: str, save_user_as: str):
        print("creating file")
        active_number = 0
        lock_number = 0
        user_active_number = 0
        user_lock_number = 0
        for data in self._split_data():
            card_no = data.card_no
            if data.is_activate:
                str_row = str(self.card_form.start + active_number)
                active_number += 1
                self._set_cell(self.card_sheet[self.card_form.ordinal + str_row], active_number)
                self._set_cell(self.card_sheet[self.card_form.card_id + str_row], data.card_id)
                self._set_cell(self.card_sheet[self.card_form.card_no + str_row], card_no)
                if data.is_month:
                    self._set_cell(self.card_sheet[self.card_form.card_type + str_row], 2)
                    if data.name:
                        str_row_user = str(self.user_form.start + user_active_number)
                        user_active_number += 1
                        self._set_cell(self.user_sheet[self.user_form.ordinal + str_row_user], user_active_number)
                        self._set_cell(self.user_sheet[self.user_form.name + str_row_user], data.name)
                        self._set_cell(self.user_sheet[self.user_form.user_id + str_row_user], data.user_id)
                        self._set_cell(self.user_sheet[self.user_form.address + str_row_user], data.address)
                        self._set_cell(self.user_sheet[self.user_form.vehicle + str_row_user], data.vehicle)
                        self._set_cell(self.user_sheet[self.user_form.end_time + str_row_user], data.end_time)
                        self._set_cell(self.user_sheet[self.user_form.card_no + str_row_user], card_no)
                else:
                    self._set_cell(self.card_sheet[self.card_form.card_type + str_row], 1)
                    self._set_cell(self.card_sheet[self.card_form.vehicle + str_row], data.vehicle)
            else:
                str_row = str(self.card_form.start + lock_number)
                lock_number += 1
                self._set_cell(self.card_sheet_lock[self.card_form.ordinal + str_row], lock_number)
                self._set_cell(self.card_sheet_lock[self.card_form.card_id + str_row], data.card_id)
                self._set_cell(self.card_sheet_lock[self.card_form.card_no + str_row], card_no)
                if data.is_month:
                    self._set_cell(self.card_sheet_lock[self.card_form.card_type + str_row], 2)
                    if data.name:
                        str_row_user = str(self.user_form.start + user_lock_number)
                        user_lock_number += 1
                        self._set_cell(self.user_sheet[self.user_form.ordinal + str_row_user], user_lock_number)
                        self._set_cell(self.user_sheet[self.user_form.name + str_row_user], data.name)
                        self._set_cell(self.user_sheet[self.user_form.user_id + str_row_user], data.user_id)
                        self._set_cell(self.user_sheet[self.user_form.address + str_row_user], data.address)
                        self._set_cell(self.user_sheet[self.user_form.vehicle + str_row_user], data.vehicle)
                        self._set_cell(self.user_sheet[self.user_form.end_time + str_row_user], data.end_time)
                        self._set_cell(self.user_sheet[self.user_form.card_no + str_row_user], card_no)
                else:
                    self._set_cell(self.card_sheet_lock[self.card_form.card_type + str_row], 1)
                    self._set_cell(self.card_sheet_lock[self.card_form.vehicle + str_row], data.vehicle)
        print("saving file")
        self.card_file.save(save_card_as)
        tail = os.path.splitext(save_card_as)[1]
        self.card_file_lock.save(save_card_as.replace(tail, "_lock" + tail))
        self.card_user.save(save_user_as)
        tail = os.path.splitext(save_user_as)[1]
        self.card_user_lock.save(save_user_as.replace(tail, "_lock" + tail))
        print("successfully")


    def _set_cell(self, cell: Cell, value: str|int, bold: bool = False, font_color: str = "00000000", size: int = 12, alignment: str = "center"):
        font = Font(bold=bold, color=font_color, size=size)
        alig = Alignment(horizontal=alignment)
        border_type = Side(border_style=borders.BORDER_THIN)
        bor = Border(top=border_type, right=border_type, bottom=border_type, left=border_type)
        cell.font = font
        cell.alignment = alig
        #cell.border = bor
        cell.value = value