"""Report generator for spring analysis

Produces a Markdown report summarizing key metrics and embedding references
to generated plots, aligned approximately with 결과보고서양식.png structure.
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import Dict, Any, Optional
import os


@dataclass
class ReportInputs:
    metrics_summary: Dict[str, float]
    basic_summary: Dict[str, float]
    params: Dict[str, Any]
    output_dir: str


def generate_report_md(path: str, inputs: ReportInputs) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    ms = inputs.metrics_summary
    bs = inputs.basic_summary
    p = inputs.params
    img1 = os.path.join(inputs.output_dir, 'cylindrical_3d.png')
    img2 = os.path.join(inputs.output_dir, 'spring_detailed_analysis.png')

    with open(path, 'w', encoding='utf-8') as f:
        f.write('# Spring Analysis Report\n\n')
        f.write('## Overview\n\n')
        f.write('- Normalization: {}\n'.format(p.get('NORMALIZATION_METHOD')))
        f.write('- SEAM mode: {}\n'.format(p.get('SEAM_MODE')))
        f.write('- THETA mode: {} (positive={})\n'.format(p.get('THETA_MODE'), p.get('THETA_POSITIVE')))
        f.write('- Arc plane: {}\n'.format(p.get('ARC_PLANE')))
        f.write('- Start strategy: {}\n'.format(p.get('START_STRATEGY')))
        f.write('- Smooth mode: {}\n'.format(p.get('SMOOTH_MODE')))
        f.write('\n')

        f.write('## Key Metrics\n\n')
        def wline(label, key):
            val = ms.get(key)
            if val is None:
                return
            try:
                f.write(f'- {label}: {val:.3f}\n')
            except Exception:
                f.write(f'- {label}: {val}\n')
        wline('Height Total', 'Height_Total')
        wline('Turn Total', 'Turn_Total')
        wline('Pitch Mean (smooth)', 'Pitch_Mean')
        wline('Pitch Min (smooth)', 'Pitch_Min')

        f.write('\n## Plots\n\n')
        if os.path.exists(img1):
            f.write(f'![Cylindrical 3D]({img1})\n\n')
        if os.path.exists(img2):
            f.write(f'![Detailed Views]({img2})\n\n')

        f.write('## Outlier & Correction Summary\n\n')
        outlier_count = p.get('OUTLIER_COUNT')
        outlier_indices = p.get('OUTLIER_INDICES')
        if outlier_count is not None:
            f.write(f'- Outlier points corrected: {outlier_count}\n')
        if outlier_indices:
            # show first few indices only
            shown = outlier_indices[:10]
            more = ' ...' if len(outlier_indices) > 10 else ''
            f.write(f'- Corrected indices (first 10): {shown}{more}\n')
        method = 'MAD radius/step with end-factor' if p.get('OUTLIER_METHOD') else 'None'
        f.write(f'- Correction method: {method}\n')
        f.write('\n## Notes\n\n')
        f.write('- This report is auto-generated based on current parameters and standardized 1000-point curve.\n')
        f.write('- Vectorial diameter and local pitch may be refined via 3-point arc interpolation ' \
                '(formula-based in Excel).\n')
