"""
IGES section-based parser (minimal for this project):
- Reads S/G/D/P/T sections
- Extracts: parameter and record delimiters, units, transformation matrices (124), points (116)
- Applies transforms and normalizes units to mm

Note: This is a lightweight implementation targeting typical ASCII IGES files
exported by Open CASCADE (observed in TK1_FRT_zero-1_251014.igs). It is more
robust than regex-only approaches while keeping complexity modest.
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional
import math
import re

@dataclass
class IgesGlobal:
    param_delim: str = ','
    record_delim: str = ';'
    units_name: str = 'MM'  # e.g., 'MM', 'IN'
    units_flag: int = 2     # 1=inches, 2=mm (typical mapping)
    scale_to_mm: float = 1.0

@dataclass
class Transform124:
    # 3D affine stored as 4x4 row-major
    m: List[List[float]]  # 4x4

    def apply(self, x: float, y: float, z: float) -> Tuple[float, float, float]:
        vx = self.m[0][0]*x + self.m[0][1]*y + self.m[0][2]*z + self.m[0][3]
        vy = self.m[1][0]*x + self.m[1][1]*y + self.m[1][2]*z + self.m[1][3]
        vz = self.m[2][0]*x + self.m[2][1]*y + self.m[2][2]*z + self.m[2][3]
        return vx, vy, vz

@dataclass
class Point116:
    x: float
    y: float
    z: float
    de_index: int  # directory entry index (1-based)
    xform_ptr: int # transform pointer from DE (0 if none)

@dataclass
class DirectoryEntry:
    entity_type: int
    param_pointer: int
    structure: int
    line_font: int
    level: int
    view: int
    xform_ptr: int
    label_assoc: int
    status: int
    weight: int
    color: int
    param_line_count: int
    form: int
    label: str
    subs: int


def _parse_global_delimiters(lines: List[str]) -> Tuple[str, str]:
    # IGES typically encodes delimiters on the first G line, like ",;"
    # Fallback to defaults if not found.
    for ln in lines:
        if ln.strip().endswith('G0000001') or ln.strip().startswith(',,'):
            # Scan for pattern like ",;"
            m = re.search(r"([,;])\s*([,;])", ln[:72])
            if m:
                return m.group(1), m.group(2)
    return ',', ';'


def _parse_global_units(lines: List[str]) -> Tuple[str, int]:
    # Try to detect tokens like "2HMM" or units flag sequence.
    text = ''.join([ln[:72] for ln in lines])
    # Common Open CASCADE pattern includes "2HMM" when millimeters
    if '2HMM' in text.upper():
        return 'MM', 2
    # Fallback simple detection
    if 'INCH' in text.upper():
        return 'IN', 1
    return 'MM', 2


def _unit_scale_to_mm(units_name: str, units_flag: int) -> float:
    name = (units_name or '').upper()
    if name in ('MM', 'MILLIMETER', 'MILLIMETRE') or units_flag == 2:
        return 1.0
    if name in ('IN', 'INCH', 'INCHES') or units_flag == 1:
        return 25.4
    return 1.0


def _chunks_80(fp: str) -> Tuple[List[str], List[str], List[str], List[str], List[str]]:
    S: List[str] = []
    G: List[str] = []
    D: List[str] = []
    P: List[str] = []
    T: List[str] = []
    with open(fp, 'r', encoding='utf-8', errors='ignore') as f:
        for ln in f:
            if len(ln) < 80:
                ln = ln.rstrip('\n')
                ln = ln + ' '*(80-len(ln))
            sec = ln[72]
            if sec == 'S':
                S.append(ln)
            elif sec == 'G':
                G.append(ln)
            elif sec == 'D':
                D.append(ln)
            elif sec == 'P':
                P.append(ln)
            elif sec == 'T':
                T.append(ln)
    return S, G, D, P, T


def _parse_directory(D: List[str]) -> List[DirectoryEntry]:
    entries: List[DirectoryEntry] = []
    # Each DE occupies two 80-char records
    for i in range(0, len(D), 2):
        a = D[i][:72]
        b = D[i+1][:72] if i+1 < len(D) else ' '*72
        # Split into 9+9 = 18 raw fields (we only retain data columns, sequence numbers already stripped).
        fields = [a[j:j+8] for j in range(0, 72, 8)] + [b[j:j+8] for j in range(0, 72, 8)]
        # Pad to 20 to avoid index errors on sparse/short records; missing tail fields set empty.
        if len(fields) < 20:
            fields.extend([''] * (20 - len(fields)))
        # Safe int parse helper
        def to_int(s: str) -> int:
            s = s.strip()
            return int(s) if s else 0
        entity_type = to_int(fields[0])
        param_pointer = to_int(fields[1])
        structure = to_int(fields[2])
        line_font = to_int(fields[3])
        level = to_int(fields[4])
        view = to_int(fields[5])
        xform_ptr = to_int(fields[6])
        label_assoc = to_int(fields[7])
        status = to_int(fields[8])
        # Second DE line indices
        weight = to_int(fields[10])
        color = to_int(fields[11])
        param_line_count = to_int(fields[12])
        form = to_int(fields[13])
        label = fields[16].strip() if len(fields) > 16 else ''
        subs = to_int(fields[17]) if len(fields) > 17 else 0
        entries.append(DirectoryEntry(
            entity_type, param_pointer, structure, line_font, level, view,
            xform_ptr, label_assoc, status, weight, color, param_line_count,
            form, label, subs
        ))
    return entries


def _tokenize_p_records(P: List[str]) -> Dict[int, str]:
    """Map P-section sequence numbers to their 72-char parameter text.
    Columns 73..80 hold the 7-digit sequence number, with 'P' at col 73.
    """
    tokens: Dict[int, str] = {}
    for ln in P:
        left = ln[:72]
        seq_str = ln[73:80]  # last 7 digits
        try:
            seq = int(seq_str)
        except ValueError:
            continue
        tokens[seq] = tokens.get(seq, '') + left
    return tokens


def _float_token(tok: str) -> float:
    t = tok.strip().replace('D', 'E').replace('d', 'e')
    # remove trailing/leading commas/spaces
    t = t.strip(', ')
    if not t:
        return 0.0
    return float(t)


def parse_iges_points(filepath: str) -> Tuple[List[Tuple[float,float,float]], IgesGlobal]:
    S, G, D, P, T = _chunks_80(filepath)
    param_delim, rec_delim = _parse_global_delimiters(G)
    units_name, units_flag = _parse_global_units(G)
    scale = _unit_scale_to_mm(units_name, units_flag)
    glb = IgesGlobal(param_delim, rec_delim, units_name, units_flag, scale)

    entries = _parse_directory(D)
    pmap = _tokenize_p_records(P)

    # Collect transforms (124)
    transforms: Dict[int, Transform124] = {}
    for idx, de in enumerate(entries, start=1):
        if de.entity_type == 124:
            # Read parameter record starting at de.param_pointer using P-sequence numbers
            raw = ''
            for k in range(de.param_pointer, de.param_pointer + max(1, de.param_line_count)):
                raw += pmap.get(k, '')
            parts = [s for s in raw.split(glb.param_delim) if s.strip()]
            # A 124 matrix encodes 12 or 16 parameters; we'll try to read 12 (3x4)
            vals: List[float] = []
            for t in parts:
                try:
                    vals.append(_float_token(t))
                except ValueError:
                    pass
            if len(vals) >= 12:
                m = [[vals[0], vals[1], vals[2], vals[3]],
                     [vals[4], vals[5], vals[6], vals[7]],
                     [vals[8], vals[9], vals[10], vals[11]],
                     [0.0, 0.0, 0.0, 1.0]]
                transforms[idx] = Transform124(m)

    # Collect points (116)
    points: List[Tuple[float,float,float]] = []
    for idx, de in enumerate(entries, start=1):
        if de.entity_type != 116:
            continue
        raw = ''
        for k in range(de.param_pointer, de.param_pointer + max(1, de.param_line_count)):
            raw += pmap.get(k, '')
        parts = [s for s in raw.split(glb.param_delim) if s.strip()]
        # Expect pattern: 116, X, Y, Z, (optional extras...)
        xyz: Optional[Tuple[float,float,float]] = None
        try:
            # Some exporters include the leading 116 in P data; handle both cases
            if parts and parts[0].strip().isdigit() and int(parts[0].strip()) == 116:
                x = _float_token(parts[1]); y = _float_token(parts[2]); z = _float_token(parts[3])
            else:
                x = _float_token(parts[0]); y = _float_token(parts[1]); z = _float_token(parts[2])
            xyz = (x, y, z)
        except Exception:
            continue
        # Apply transform if present
        if de.xform_ptr and de.xform_ptr in transforms:
            Tm = transforms[de.xform_ptr]
            xyz = Tm.apply(*xyz)
        # Unit scaling to mm
        xyz = (xyz[0] * glb.scale_to_mm, xyz[1] * glb.scale_to_mm, xyz[2] * glb.scale_to_mm)
        points.append(xyz)

    return points, glb
