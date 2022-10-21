from dataclasses import dataclass

@dataclass
class Service:
    '''service's info what it's used by Costumer
    attribute type can be "vaccine" or "deworm"
    attribute priority  can be a number between 0 to 4
    '''
    type: str
    mediccine: str
    date: str
    priority: int = 0

    def set_priority(RGB: str) -> int:
        pass


@dataclass
class Costumer:
    '''costumer's info
    '''
    name: str
    pet_name: str
    race: str
    phone: str
    vaccine: Service
    deworn: Service
