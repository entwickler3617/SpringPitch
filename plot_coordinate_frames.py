#!/usr/bin/env python3
"""
원점을 기준으로 직교좌표계(x,y,z)와 원통좌표계(r,θ,z)를 3D로 시각화합니다.
- 직교축: X(빨강), Y(초록), Z(파랑)
- 원통: r 벡터(주황), θ 호(보라), XY 평면의 원형 그리드, Z축 공유
결과 이미지는 coordinate_frames.png 로 저장됩니다.
"""
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401 (needed for 3D)
import logging
import os

# 스타일 설정
plt.rcParams["axes.unicode_minus"] = False

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def draw_cartesian_axes(ax, L=1.5):
    # 원점
    O = np.array([0.0, 0.0, 0.0])
    # X, Y, Z 축
    ax.plot([0, L], [0, 0], [0, 0], color="red", lw=2)
    ax.plot([0, 0], [0, L], [0, 0], color="green", lw=2)
    ax.plot([0, 0], [0, 0], [0, L], color="blue", lw=2)
    ax.text(L, 0, 0, "X", color="red")
    ax.text(0, L, 0, "Y", color="green")
    ax.text(0, 0, L, "Z", color="blue")


def draw_cylindrical(ax, R=1.0, theta=np.pi/4, H=1.2):
    # XY 평면의 원형 그리드 (r 고정 예시)
    t = np.linspace(0, 2*np.pi, 200)
    ax.plot(R*np.cos(t), R*np.sin(t), 0*t, color="gray", lw=1, alpha=0.7)

    # 몇 개의 반지름 가이드
    for r in np.linspace(R*0.25, R, 4):
        ax.plot(r*np.cos(t), r*np.sin(t), 0*t, color="lightgray", lw=0.8, alpha=0.6)

    # θ 호 (0 ~ theta)
    tt = np.linspace(0, theta, 100)
    ax.plot(0.6*np.cos(tt), 0.6*np.sin(tt), 0*tt, color="purple", lw=2)
    ax.text(0.65*np.cos(theta/2), 0.65*np.sin(theta/2), 0, "θ", color="purple")

    # r 벡터 (XY 평면, z=0)
    rx, ry = np.cos(theta), np.sin(theta)
    ax.plot([0, rx], [0, ry], [0, 0], color="orange", lw=3)
    ax.text(rx, ry, 0, "r", color="orange")

    # 원통 Z축(직교 Z와 동일)
    ax.plot([0, 0], [0, 0], [0, H], color="blue", lw=2, alpha=0.6)

    # 같은 각도에서의 수직 선(원통면 가이드)
    ax.plot([rx, rx], [ry, ry], [0, H], color="gray", lw=1, ls=":", alpha=0.8)


def main():
    fig = plt.figure(figsize=(8, 7))
    ax = fig.add_subplot(111, projection="3d")

    draw_cartesian_axes(ax, L=1.5)
    draw_cylindrical(ax, R=1.0, theta=np.deg2rad(45), H=1.2)

    # 보기 설정
    ax.set_xlim(-1.2, 1.8)
    ax.set_ylim(-1.2, 1.8)
    ax.set_zlim(0, 1.5)
    ax.set_xlabel("X")
    ax.set_ylabel("Y")
    ax.set_zlabel("Z")
    ax.set_title("Cartesian (X,Y,Z) and Cylindrical (r,θ,Z) Frames at Origin")
    ax.set_box_aspect((1,1,0.9))
    ax.view_init(elev=22, azim=-55)

    plt.tight_layout()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(script_dir, "coordinate_frames.png")
    plt.savefig(out_path, dpi=160, bbox_inches="tight")
    logger.info(f"Saved: {out_path}")
    plt.show()


if __name__ == "__main__":
    main()
