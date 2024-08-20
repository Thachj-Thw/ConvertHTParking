from openpyxl import Workbook, load_workbook
from openpyxl.cell import Cell
import module
from .structs import InputForm, OutputUserForm, OutputCardForm, DataStruct
from .autoformat import format_card_code, format_vehicle_plate


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
        card_no     = "B",
        card_id     = "C",
        card_type   = "D",
        vehicle     = "E",
        status      = "G",
        start       = 6
    )

    user_form = OutputUserForm(
        name        = "B",
        address     = "C",
        user_id     = "E",
        end_time    = "K",
        vehicle     = "P",
        card_no     = "W",
        start       = 5
    )

    def __init__(self, file_path: str):
        self.workbook = load_workbook(filename=file_path)
        self.sheet = self.workbook.active
        self.card_file = load_workbook(filename=path.source.join("templates", "BaseCard.xlsx"))
        self.card_user = load_workbook(filename=path.source.join("templates", "BaseUser.xlsx"))
        self.all_data = []
        print(self._split_data())
    
    def _split_data(self):
        all_data = []
        for row in range(self.input_form.start, 1498):
            card_no = self.sheet[self.input_form.card_no + str(row)].value
            card_id = self.sheet[self.input_form.card_id + str(row)].value
            vehicle = self.sheet[self.input_form.vehicle + str(row)].value
            vehicle_plate = self.sheet[self.input_form.vehicle + str(row)].value
            end_time = self.sheet[self.input_form.end_time + str(row)].value
            status = self.sheet[self.input_form.status + str(row)].value
            user_id = self.sheet[self.input_form.user_id + str(row)].value
            name = self.sheet[self.input_form.name + str(row)].value
            address = self.sheet[self.input_form.address + str(row)].value
            data = DataStruct(card_no, card_id, user_id, True if status == "Hoạt động" else False ,vehicle, vehicle_plate, end_time, name, address)
            all_data.append(data)
        return all_data
