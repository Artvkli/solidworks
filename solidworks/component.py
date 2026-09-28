from dataclasses import dataclass
from pathlib import Path


@dataclass
class Component:
    name: str
    path: Path
    component_type: str
    suppressed: bool = False

    @property
    def is_part(self):
        return self.component_type == "PART"

    @property
    def is_assembly(self):
        return self.component_type == "ASSEMBLY"    