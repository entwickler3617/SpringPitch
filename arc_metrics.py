"""3-point arc interpolation metrics (initial scaffolding)

Provides arc-based refinement of local diameter and pitch using triplets of standardized points.
Rules:
 - For each i, consider (i-1, i, i+1) triplet in cylindrical space.
 - Fit a circle through the 3 points (in XY plane of local frame) to estimate instantaneous radius.
 - Compute vectorial diameter as 2*estimated radius.
 - Estimate local pitch via dZ per one turn equivalent using local angular differences.

Note: This initial version uses planar XY circle fit; for general 3D springs aligned to PCA axis,
XY of local frame is perpendicular cross-section.
"""
from __future__ import annotations
import numpy as np
from typing import Tuple
import os


def circle_from_3pts(xy1: np.ndarray, xy2: np.ndarray, xy3: np.ndarray) -> Tuple[float, float, float]:
    """Return circle center (cx, cy) and radius r from three non-collinear points.
    If degenerate, returns (nan, nan, nan).
    """
    x1,y1 = xy1; x2,y2 = xy2; x3,y3 = xy3
    a = x1*(y2 - y3) - y1*(x2 - x3) + x2*y3 - x3*y2
    if abs(a) < 1e-12:
        return float('nan'), float('nan'), float('nan')
    b = (x1**2 + y1**2)*(y3 - y2) + (x2**2 + y2**2)*(y1 - y3) + (x3**2 + y3**2)*(y2 - y1)
    c = (x1**2 + y1**2)*(x2 - x3) + (x2**2 + y2**2)*(x3 - x1) + (x3**2 + y3**2)*(x1 - x2)
    cx = -b/(2*a)
    cy = -c/(2*a)
    r = np.hypot(x1 - cx, y1 - cy)
    return float(cx), float(cy), float(r)


def arc_refine_radius_and_pitch(local_std_xyz: np.ndarray, theta_unwrapped: np.ndarray,
                                 z: np.ndarray) -> Tuple[np.ndarray, np.ndarray]:
    """Refine diameter and pitch using 3-point arcs.
    Plane selection via env SPRING_ARC_PLANE: 'local_xy' (default) | 'best_fit'
    - best_fit: fits a local plane via PCA on triplet before circle fit (project points).
    Returns: (diameter_vectorial, pitch_arc)
    """
    N = len(theta_unwrapped)
    if N < 3:
        return np.full(N, np.nan), np.full(N, np.nan)
    mode = os.getenv('SPRING_ARC_PLANE', 'local_xy').lower()
    x = local_std_xyz[:,0]; y = local_std_xyz[:,1]
    diam_vec = np.full(N, np.nan)
    pitch_arc = np.full(N, np.nan)
    for i in range(1, N-1):
        if mode == 'best_fit':
            pts = local_std_xyz[i-1:i+2]
            # PCA plane: subtract centroid
            c = pts.mean(axis=0)
            q = pts - c
            # singular vectors
            U, S, Vt = np.linalg.svd(q, full_matrices=False)
            # project to first two principal components
            proj = q @ Vt[:2].T
            cxp, cyp, r = circle_from_3pts(proj[0], proj[1], proj[2])
            # r 그대로 사용 (직교성 보존 가정)
        else:
            cx, cy, r = circle_from_3pts(np.array([x[i-1], y[i-1]]), np.array([x[i], y[i]]), np.array([x[i+1], y[i+1]]))
        if not np.isnan(r):
            diam_vec[i] = 2.0 * r
        dtheta = theta_unwrapped[i+1] - theta_unwrapped[i-1]
        dz = z[i+1] - z[i-1]
        if abs(dtheta) > 1e-8:
            pitch_arc[i] = dz / (dtheta / (2.0*np.pi))
    # Edge handling: copy nearest valid
    if N >= 3:
        diam_vec[0] = diam_vec[1]
        diam_vec[-1] = diam_vec[-2]
        pitch_arc[0] = pitch_arc[1]
        pitch_arc[-1] = pitch_arc[-2]
    return diam_vec, pitch_arc
