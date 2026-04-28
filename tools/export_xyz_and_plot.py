import os
import sys
import csv
import traceback

# Ensure project root on path
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

from detailed_center_analysis import parse_igs_points


def ensure_dir(path: str):
    if not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)


def main():
    igs_path = os.path.join(ROOT, 'TK1_FRT_zero-1_251014.igs')
    if len(sys.argv) > 1:
        igs_path = sys.argv[1]

    out_dir = os.path.join(ROOT, 'output')
    ensure_dir(out_dir)
    csv_path = os.path.join(out_dir, 'points_xyz.csv')
    img_path = os.path.join(out_dir, 'points_preview.png')

    print(f"IGS file: {igs_path}")
    if not os.path.exists(igs_path):
        print("ERROR: IGS file not found.")
        sys.exit(2)

    try:
        pts = parse_igs_points(igs_path)
        n = len(pts)
        print(f"Parsed points: {n}")

        # Save CSV
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f)
            w.writerow(['x','y','z'])
            for p in pts:
                # p may be numpy array
                w.writerow([float(p[0]), float(p[1]), float(p[2])])
        print(f"CSV saved: {csv_path}")

        # Plot 3D line (sequential)
        fig = plt.figure(figsize=(8, 6))
        ax = fig.add_subplot(111, projection='3d')
        ax.plot(pts[:,0], pts[:,1], pts[:,2], linewidth=1)
        ax.set_title('IGS Points Preview (Sequential)')
        ax.set_xlabel('X (mm)')
        ax.set_ylabel('Y (mm)')
        ax.set_zlabel('Z (mm)')
        # Autoscale and save
        plt.tight_layout()
        plt.savefig(img_path, dpi=200, bbox_inches='tight')
        plt.close(fig)
        print(f"Preview saved: {img_path}")

        return 0
    except Exception:
        print("Failed to export/plot IGS points:")
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
