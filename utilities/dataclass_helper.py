from dataclasses import dataclass


@dataclass(unsafe_hash=True)
class Factory:
    code: str
    name: str
    address: str
    total: str
    signatory: str
