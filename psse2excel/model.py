"""ParsedData: the parsed tables of one raw file."""

import datetime
from dataclasses import dataclass, field
from typing import Any, Dict, List, NamedTuple, Optional, Tuple

import pandas as pd

from . import layouts as L


class Tab(NamedTuple):
    attr: str           # ParsedData attribute
    sheet: str          # Excel tab name
    description: str
    columns: List[str]  # column order (used for the header when there are no records)


_TWO_W = L.columns(L.TWO_WINDING)
_THREE_W = L.columns(L.THREE_WINDING)

TABS: List[Tab] = [
    Tab('system_wide', 'System_Wide_Data', 'System-wide data (solution parameters, rating sets)',
        L.SYSTEM_WIDE_COLUMNS),
    Tab('buses', 'Buses', 'Bus data with voltage levels and types', L.columns(L.BUS)),
    Tab('loads', 'Loads', 'Load data with real and reactive power', L.columns(L.LOAD)),
    Tab('fixed_shunts', 'Fixed_Shunts', 'Fixed shunt compensation devices',
        L.columns(L.FIXED_SHUNT)),
    Tab('generators', 'Generators', 'Generator data with limits and settings',
        L.columns(L.GENERATOR)),
    Tab('branches', 'Branches', 'Transmission line data', L.columns(L.BRANCH)),
    Tab('system_switching_devices', 'System_Switching_Devices',
        'System switching devices (breakers and switches)',
        L.columns(L.SYSTEM_SWITCHING_DEVICE)),
    Tab('two_winding_transformers', '2W_Transformers', 'Two-winding transformer data', _TWO_W),
    Tab('three_winding_transformers', '3W_Transformers', 'Three-winding transformer data',
        _THREE_W),
    Tab('areas', 'Areas', 'Area interchange data', L.columns(L.AREA)),
    Tab('two_terminal_dc_lines', 'Two_Terminal_DC_Lines', 'Two terminal DC lines',
        L.columns(L.TWO_TERMINAL_DC)),
    Tab('vsc_dc_lines', 'VSC_DC_lines', 'VSC DC lines', L.columns(L.VSC_DC)),
    Tab('impedance_correction', 'Impedance_Correction', 'Impedance correction data',
        ['I', 'T1', 'Re(F1)', 'Im(F1)']),
    Tab('mtdc_lines', 'Multi_Terminal_DC_Lines', 'Multi-terminal DC lines',
        L.columns(L.MTDC_LINE)),
    Tab('mtdc_converters', 'Multi_Terminal_DC_Converters', 'Multi-terminal DC converters',
        ['NAME'] + L.columns(L.MTDC_CONVERTER)),
    Tab('mtdc_dc_buses', 'Multi_Terminal_DC_Buses', 'Multi-terminal DC buses',
        ['NAME'] + L.columns(L.MTDC_DC_BUS)),
    Tab('mtdc_dc_links', 'Multi_Terminal_DC_Links', 'Multi-terminal DC links',
        ['NAME'] + L.columns(L.MTDC_DC_LINK)),
    Tab('multi_section_lines', 'Multi_Section_Lines', 'Multi-section line data',
        L.columns(L.MULTI_SECTION_LINE)),
    Tab('zones', 'Zones', 'Zone data', L.columns(L.ZONE)),
    Tab('inter_area_transfers', 'Inter_Area_Transfers', 'Inter-area transfer data',
        L.columns(L.INTER_AREA_TRANSFER)),
    Tab('owners', 'Owners', 'Owner data', L.columns(L.OWNER)),
    Tab('facts', 'FACTS', 'FACTS device data', L.columns(L.FACTS)),
    Tab('switched_shunts', 'Switched_Shunts', 'Switched shunt compensation devices',
        L.columns(L.SWITCHED_SHUNT)),
    Tab('gne_devices', 'GNE_Devices', 'Generic Network Element devices', L.columns(L.GNE)),
    Tab('induction_machines', 'Induction_Machines', 'Induction machine data',
        L.columns(L.INDUCTION_MACHINE)),
    Tab('substations', 'Substations', 'Substation data', L.columns(L.SUBSTATION)),
    Tab('substation_nodes', 'Substation_Nodes', 'Substation node data',
        ['IS'] + L.columns(L.SUBSTATION_NODE)),
    Tab('substation_switching_devices', 'Substation_Switching_Devices',
        'Substation switching device data', ['IS'] + L.columns(L.SUBSTATION_SWITCHING_DEVICE)),
    Tab('substation_connections', 'Substation_Connections',
        'Substation equipment connection data', ['IS'] + L.columns(L.SUBSTATION_CONNECTION)),
]
TAB_BY_ATTR = {t.attr: t for t in TABS}


def _frames():
    return field(default_factory=pd.DataFrame)


@dataclass
class ParsedData:
    """Container for the parsed sections of one raw file, one DataFrame per tab."""
    title: str = ""
    base_mva: float = 100.0
    version: int = 35          # the version the file was read as
    file_version: Optional[int] = None   # REV from the file's header (None if missing)
    source_file: str = ""
    # records that could not be read, per tab, and lines that belong to no record
    skipped: Dict[str, int] = field(default_factory=dict)
    lines_not_read: int = 0

    system_wide: pd.DataFrame = _frames()
    buses: pd.DataFrame = _frames()
    loads: pd.DataFrame = _frames()
    fixed_shunts: pd.DataFrame = _frames()
    generators: pd.DataFrame = _frames()
    branches: pd.DataFrame = _frames()
    system_switching_devices: pd.DataFrame = _frames()
    two_winding_transformers: pd.DataFrame = _frames()
    three_winding_transformers: pd.DataFrame = _frames()
    areas: pd.DataFrame = _frames()
    two_terminal_dc_lines: pd.DataFrame = _frames()
    vsc_dc_lines: pd.DataFrame = _frames()
    impedance_correction: pd.DataFrame = _frames()
    mtdc_lines: pd.DataFrame = _frames()
    mtdc_converters: pd.DataFrame = _frames()
    mtdc_dc_buses: pd.DataFrame = _frames()
    mtdc_dc_links: pd.DataFrame = _frames()
    multi_section_lines: pd.DataFrame = _frames()
    zones: pd.DataFrame = _frames()
    inter_area_transfers: pd.DataFrame = _frames()
    owners: pd.DataFrame = _frames()
    facts: pd.DataFrame = _frames()
    switched_shunts: pd.DataFrame = _frames()
    gne_devices: pd.DataFrame = _frames()
    induction_machines: pd.DataFrame = _frames()
    substations: pd.DataFrame = _frames()
    substation_nodes: pd.DataFrame = _frames()
    substation_switching_devices: pd.DataFrame = _frames()
    substation_connections: pd.DataFrame = _frames()

    @property
    def total_skipped(self) -> int:
        return sum(self.skipped.values())

    def to_excel(self, filename: str, include_summary: bool = True) -> None:
        """Write one tab per section (plus Summary and System_Info) to an Excel file."""
        from .excel import write_workbook
        write_workbook(self, filename, include_summary=include_summary)

    def to_csv(self, folder: str, include_summary: bool = True) -> List[str]:
        """Write one CSV file per tab (plus Summary and System_Info) into a folder.

        Returns the paths written."""
        from .csv_export import write_csv_folder
        return write_csv_folder(self, folder, include_summary=include_summary)

    def version_text(self):
        """The version for System_Info: e.g. 35, or '36 (read as 35)' for an unsupported one."""
        if self.file_version == self.version:
            return self.version
        found = 'unknown' if self.file_version is None else self.file_version
        return f'{found} (read as {self.version})'

    def summary_table(self) -> Tuple[List[str], List[list]]:
        """The Summary tab: one row per tab."""
        columns = ['Section', 'Tab', 'Record_Count', 'Skipped_Records', 'Description']
        rows = [[t.attr, t.sheet, len(getattr(self, t.attr)), self.skipped.get(t.attr, 0),
                 t.description] for t in TABS]
        return columns, rows

    def system_info_table(self) -> Tuple[List[str], List[list]]:
        """The System_Info tab: one row per parameter."""
        rows = [
            ['System Title', self.title],
            ['Base MVA', self.base_mva],
            ['Total Buses', len(self.buses)],
            ['Total Generators', len(self.generators)],
            ['Total Loads', len(self.loads)],
            ['Source File', self.source_file],
            ['PSS/E Version', self.version_text()],
            ['Run Date', datetime.datetime.now().strftime('%Y-%m-%d %H:%M')],
            ['Records Skipped', self.total_skipped],
            ['Lines Not Read', self.lines_not_read],
        ]
        return ['Parameter', 'Value'], rows

    def get_statistics(self) -> Dict[str, Any]:
        """Record counts and a few totals."""
        stats: Dict[str, Any] = {
            'system_info': {'title': self.title, 'base_mva': self.base_mva,
                            'version': self.version},
            'record_counts': {t.attr: len(getattr(self, t.attr)) for t in TABS},
        }
        if not self.buses.empty:
            stats['voltage_levels'] = self.buses['BASE_KV'].value_counts().to_dict()
        if not self.generators.empty:
            stats['generation_stats'] = {
                'total_capacity_mw': self.generators['PMAX'].sum(),
                'total_generation_mw': self.generators['PG'].sum(),
                'average_generator_size_mw': self.generators['PMAX'].mean(),
            }
        if not self.loads.empty:
            stats['load_stats'] = {
                'total_load_mw': self.loads['PL'].sum(),
                'total_load_mvar': self.loads['QL'].sum(),
            }
        return stats
