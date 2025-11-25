from enum import StrEnum

class WorkerTypes(StrEnum):
    Fest = "Fest"
    Mini = "Mini"
    Gewerblich = "Gewerblich"

class TravelStatus(StrEnum):
    Anreise = "Anreise"
    Abreise = "Abreise"
    Away24h = "24h_away"