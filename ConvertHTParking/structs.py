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
    ordinal: str
    card_no: str
    card_id: str
    card_type: str
    vehicle: str
    status: str
    start: int


@dataclass
class OutputUserForm:
    ordinal: str
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
    is_month: bool

    def is_emty(self) -> bool:
        return not self.card_no and not self.card_id and not self.user_id \
            and not self.vehicle and not self.vehicle_plate and not self.end_time and not self.name and not self.address
