#!/usr/bin/env python3
"""
Final Report Chart Generator for Spring Analysis

Creates a comprehensive summary chart with key metrics and visualizations.
"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import logging
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_final_report_chart(excel_path, output_path):
    """Create a comprehensive final report chart from TK1.xlsx data."""

    # Read the data
    df = pd.read_excel(excel_path, sheet_name='zero-1')

    # Extract key columns
    turns = df['N_turn'].values.astype(float)
    radius = df['O_radius_copy'].values.astype(float)
    pitch = df['W_pitch'].values.astype(float)
    height = df['U_rel_height'].values.astype(float)
    theta_deg = df['θ'].values.astype(float)
    z_abs = df['T_abs_z'].values.astype(float)
    x_std = df['x_norm'].values.astype(float)
    y_std = df['y_norm'].values.astype(float)
    z_std = df['z_norm'].values.astype(float)

    # Calculate summary metrics
    total_height = np.max(height) - np.min(height)
    mean_pitch = np.nanmean(pitch[pitch > 0])
    min_pitch = np.nanmin(pitch[pitch > 0])
    total_turns = np.max(turns) - np.min(turns)
    mean_radius = np.mean(radius)
    max_radius = np.max(radius)
    min_radius = np.min(radius)

    # Create figure with subplots
    fig = plt.figure(figsize=(16, 12))
    fig.suptitle('Spring Analysis Final Report', fontsize=16, fontweight='bold')

    # 1. 3D Spring Visualization
    ax1 = fig.add_subplot(2, 3, 1, projection='3d')
    ax1.plot(x_std, y_std, z_std, 'b-', linewidth=1, alpha=0.7)
    ax1.scatter(x_std[0], y_std[0], z_std[0], c='red', s=50, label='Start')
    ax1.scatter(x_std[-1], y_std[-1], z_std[-1], c='green', s=50, label='End')
    ax1.set_xlabel('X (mm)')
    ax1.set_ylabel('Y (mm)')
    ax1.set_zlabel('Z (mm)')
    ax1.set_title('3D Spring Geometry')
    ax1.legend()
    ax1.grid(True)

    # 2. Key Metrics Summary (Text Box)
    ax2 = fig.add_subplot(2, 3, 2)
    ax2.axis('off')
    metrics_text = f"""
    SPRING ANALYSIS SUMMARY

    Dimensions:
    • Total Height: {total_height:.2f} mm
    • Total Turns: {total_turns:.2f}
    • Mean Radius: {mean_radius:.2f} mm
    • Radius Range: {min_radius:.2f} - {max_radius:.2f} mm

    Pitch Analysis:
    • Mean Pitch: {mean_pitch:.2f} mm
    • Min Pitch: {min_pitch:.2f} mm

    Data Points: {len(df)} (standardized)
    """
    ax2.text(0.1, 0.9, metrics_text, transform=ax2.transAxes,
             fontsize=10, verticalalignment='top', fontfamily='monospace',
             bbox=dict(boxstyle='round', facecolor='lightblue', alpha=0.5))

    # 3. Radius vs Turns
    ax3 = fig.add_subplot(2, 3, 3)
    ax3.plot(turns, radius, 'r-', linewidth=1.5)
    ax3.set_xlabel('Turns')
    ax3.set_ylabel('Radius (mm)')
    ax3.set_title('Radius vs Turns')
    ax3.grid(True)
    ax3.fill_between(turns, radius, alpha=0.3, color='red')

    # 4. Pitch vs Turns
    ax4 = fig.add_subplot(2, 3, 4)
    valid_pitch = pitch > 0
    ax4.plot(turns[valid_pitch], pitch[valid_pitch], 'g-', linewidth=1.5, marker='o', markersize=2)
    ax4.set_xlabel('Turns')
    ax4.set_ylabel('Pitch (mm)')
    ax4.set_title('Pitch vs Turns')
    ax4.grid(True)
    ax4.axhline(y=float(mean_pitch), color='orange', linestyle='--', label=f'Mean: {mean_pitch:.2f} mm')
    ax4.legend()

    # 5. Theta vs Height
    ax5 = fig.add_subplot(2, 3, 5)
    ax5.plot(height, theta_deg, 'purple', linewidth=1.5)
    ax5.set_xlabel('Height (mm)')
    ax5.set_ylabel('Theta (degrees)')
    ax5.set_title('Theta vs Height')
    ax5.grid(True)

    # 6. Radius Distribution Histogram
    ax6 = fig.add_subplot(2, 3, 6)
    ax6.hist(radius, bins=30, alpha=0.7, color='skyblue', edgecolor='black')
    ax6.axvline(x=float(mean_radius), color='red', linestyle='--', linewidth=2,
                label=f'Mean: {mean_radius:.2f} mm')
    ax6.set_xlabel('Radius (mm)')
    ax6.set_ylabel('Frequency')
    ax6.set_title('Radius Distribution')
    ax6.legend()
    ax6.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

    logger.info(f"Final report chart saved to: {output_path}")

if __name__ == "__main__":
    # Paths - use relative paths for portability
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(script_dir, "output", "TK1.xlsx")
    output_path = os.path.join(script_dir, "output", "final_report.png")

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Check if Excel file exists
    if not os.path.exists(excel_path):
        logger.error(f"Error: {excel_path} not found. Please run the analysis first.")
        exit(1)

    # Generate the report
    create_final_report_chart(excel_path, output_path)