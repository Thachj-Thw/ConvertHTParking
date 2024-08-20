from dataclasses import dataclass, field


@dataclass
class InputForm:
    card_no: str
    card_id: str
    user_id: str
    vehicle: str
    vehicle_plate: str
    end_time: str
    status: str
    name: str
    address: str
    start: int


@dataclass
class OutputCardForm:
    card_no: str
    card_id: str
    card_type: str
    vehicle: str
    status: str
    start: int


@dataclass
class OutputUserForm:
    card_no: str
    user_id: str
    vehicle: str
    name: str
    address: str
    end_time: str
    start: int



@dataclass
class DataStruct:
    card_no: str
    card_id: str
    user_id: str
    is_activate: bool
    vehicle: str
    vehicle_plate: str
    end_time: str
    name: str
    address: str
