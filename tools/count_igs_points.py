import os
import sys
import traceback

# Ensure the project root is on sys.path
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from detailed_center_analysis import parse_igs_points


def main():
    # Default igs path in workspace
    igs_path = os.path.join(ROOT, 'TK1_FRT_zero-1_251014.igs')
    if len(sys.argv) > 1:
        igs_path = sys.argv[1]

    print(f"IGS file: {igs_path}")
    if not os.path.exists(igs_path):
        print("ERROR: IGS file not found.")
        sys.exit(2)

    try:
        pts = parse_igs_points(igs_path)
        n = len(pts)
        print(f"Point count: {n}")
        # Print a quick preview of first 3 points
        for i, p in enumerate(pts[:3]):
            print(f"P{i}: {p.tolist()}")
        sys.exit(0)
    except Exception as e:
        print("Failed to parse IGS file:")
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
