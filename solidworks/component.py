from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, List, Optional


@dataclass
class Component:
    name: str
    path: Path
    component_type: str
    suppressed: bool = False
    sw_component: Any = None
    lightweight: bool = False
    configuration: Optional[str] = None

    @property
    def is_part(self) -> bool:
        return self.component_type.upper() == "PART"

    @property
    def is_assembly(self) -> bool:
        return self.component_type.upper() == "ASSEMBLY"


@dataclass
class PartRecord:
    name: str
    path: Path
    quantity: int = 1
    configuration: Optional[str] = None
    components: List[Component] = field(default_factory=list)
