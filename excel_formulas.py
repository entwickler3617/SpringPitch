"""Excel formulas replication module

Implements computations corresponding to TK1.xlsx std sheet L~X columns and Z3:AC14 summary block.
Focus metrics (initial version):
 - Turn (already computed upstream) -> pass-through
 - Radius, Diameter(Perpendicular)=2*R
 - Diameter(Vectorial) (placeholder = 2*R; arc-based refinement later)
 - Height incremental and total
 - Local pitch and min pitch (difference-based + optional arc refinement placeholder)

Functions here expect standardized arrays (length N):
    r: radius values
    z: axial coordinate values
    theta_unwrapped: continuous theta (rad)

Returned dict fields align with column naming expected for zero-1 sheet integration.
"""
from __future__ import annotations
from dataclasses import dataclass
import numpy as np
from typing import Dict, Any

@dataclass
class BasicMetrics:
    turn: np.ndarray
    radius: np.ndarray
    diameter_perp: np.ndarray
    diameter_vectorial: np.ndarray
    height_rel: np.ndarray
    pitch_local: np.ndarray
    pitch_local_smooth: np.ndarray
    pitch_min: float
    pitch_mean: float
    height_total: float


def moving_average(x: np.ndarray, window: int = 9) -> np.ndarray:
    if window < 2:
        return x.copy()
    if window % 2 == 0:
        window += 1
    pad = window // 2
    xpad = np.pad(x, (pad, pad), mode='edge')
    k = np.ones(window)/window
    y = np.convolve(xpad, k, mode='same')[pad:-pad]
    return y


def compute_basic_metrics_from_cyl(r: np.ndarray, z: np.ndarray, theta_unwrapped: np.ndarray) -> BasicMetrics:
    N = len(r)
    if N == 0:
        return BasicMetrics(
            turn=np.array([]), radius=r, diameter_perp=r, diameter_vectorial=r,
            height_rel=z, pitch_local=np.array([]), pitch_local_smooth=np.array([]),
            pitch_min=float('nan'), pitch_mean=float('nan'), height_total=float('nan')
        )
    turn = theta_unwrapped / (2.0 * np.pi)
    # Local pitch dZ/dTurn
    dturn = np.diff(turn)
    dz = np.diff(z)
    with np.errstate(divide='ignore', invalid='ignore'):
        local_pitch_seg = np.where(np.abs(dturn) > 1e-10, dz / dturn, np.nan)
    pitch_local = np.empty(N)
    pitch_local[0] = local_pitch_seg[0] if len(local_pitch_seg) else np.nan
    pitch_local[1:] = local_pitch_seg
    # Smooth (window heuristic: up to ~N/50, odd)
    w = max(5, (N // 50) * 2 + 1)
    w = min(w, 101)
    if w % 2 == 0: w += 1
    pitch_smooth = moving_average(pitch_local, w)
    diameter_perp = 2.0 * r
    diameter_vec = diameter_perp.copy()  # placeholder; arc refinement later
    height_rel = z - z[0]
    height_total = float(z.max() - z.min()) if N > 0 else float('nan')
    pitch_min = float(np.nanmin(pitch_smooth)) if np.any(~np.isnan(pitch_smooth)) else float('nan')
    pitch_mean = float(np.nanmean(pitch_smooth)) if np.any(~np.isnan(pitch_smooth)) else float('nan')
    return BasicMetrics(
        turn=turn,
        radius=r,
        diameter_perp=diameter_perp,
        diameter_vectorial=diameter_vec,
        height_rel=height_rel,
        pitch_local=pitch_local,
        pitch_local_smooth=pitch_smooth,
        pitch_min=pitch_min,
        pitch_mean=pitch_mean,
        height_total=height_total,
    )


def metrics_to_columns(metrics: BasicMetrics) -> Dict[str, Any]:
    return {
        'Turn': metrics.turn,
        'Radius': metrics.radius,
        'Diameter_Perpendicular': metrics.diameter_perp,
        'Diameter_Vectorial': metrics.diameter_vectorial,
        'Height_Rel': metrics.height_rel,
        'Pitch_Local': metrics.pitch_local,
        'Pitch_Smooth': metrics.pitch_local_smooth,
    }


def summary_from_metrics(metrics: BasicMetrics) -> Dict[str, float]:
    return {
        'Pitch_Min': metrics.pitch_min,
        'Pitch_Mean': metrics.pitch_mean,
        'Height_Total': metrics.height_total,
        'Turn_Total': float(metrics.turn[-1] - metrics.turn[0]) if len(metrics.turn) else float('nan')
    }
