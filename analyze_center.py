#!/usr/bin/env python3
"""
IGS 파일로부터 스프링 데이터의 중심점을 분석하는 스크립트
"""

import re
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import logging
import os

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def parse_igs_points(filepath):
    """IGS 파일에서 좌표점들을 추출"""
    points = []
    
    # 좌표 데이터를 찾는 정규식 패턴
    pattern = r'116,(-?\d+\.?\d*),(-?\d+\.?\d*),(-?\d+\.?\d*),0;'
    
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            for line in file:
                match = re.search(pattern, line)
                if match:
                    x = float(match.group(1))
                    y = float(match.group(2))
                    z = float(match.group(3))
                    points.append([x, y, z])
    except Exception as e:
        logger.error(f"파일 읽기 오류: {e}")
        return None
    
    return np.array(points)

def analyze_spring_center(points):
    """스프링 데이터의 중심과 특성 분석"""
    if points is None or len(points) == 0:
        return None
    
    logger.info(f"총 데이터 포인트 수: {len(points)}")
    
    # 기본 통계
    min_vals = np.min(points, axis=0)
    max_vals = np.max(points, axis=0)
    mean_vals = np.mean(points, axis=0)
    std_vals = np.std(points, axis=0)
    
    logger.info("=== XYZ 좌표 분석 ===")
    logger.info(f"X 범위: {min_vals[0]:.3f} ~ {max_vals[0]:.3f}")
    logger.info(f"Y 범위: {min_vals[1]:.3f} ~ {max_vals[1]:.3f}") 
    logger.info(f"Z 범위: {min_vals[2]:.3f} ~ {max_vals[2]:.3f}")
    logger.info(f"X 중심: {mean_vals[0]:.3f} (표준편차: {std_vals[0]:.3f})")
    logger.info(f"Y 중심: {mean_vals[1]:.3f} (표준편차: {std_vals[1]:.3f})")
    logger.info(f"Z 중심: {mean_vals[2]:.3f} (표준편차: {std_vals[2]:.3f})")
    
    # 스프링 축 분석 (Z축 방향 가정)
    z_min_idx = np.argmin(points[:, 2])
    z_max_idx = np.argmax(points[:, 2])
    
    spring_bottom = points[z_min_idx]
    spring_top = points[z_max_idx]
    spring_height = max_vals[2] - min_vals[2]
    
    logger.info("=== 스프링 구조 분석 ===")
    logger.info(f"스프링 하단점: ({spring_bottom[0]:.3f}, {spring_bottom[1]:.3f}, {spring_bottom[2]:.3f})")
    logger.info(f"스프링 상단점: ({spring_top[0]:.3f}, {spring_top[1]:.3f}, {spring_top[2]:.3f})")
    logger.info(f"스프링 높이: {spring_height:.3f}")
    
    # XY 평면에서의 반지름 분석
    center_xy = mean_vals[:2]  # XY 평면 중심
    radii = np.sqrt((points[:, 0] - center_xy[0])**2 + (points[:, 1] - center_xy[1])**2)
    
    logger.info("=== 반지름 분석 (XY 평면) ===")
    logger.info(f"평균 반지름: {np.mean(radii):.3f}")
    logger.info(f"최소 반지름: {np.min(radii):.3f}")
    logger.info(f"최대 반지름: {np.max(radii):.3f}")
    logger.info(f"반지름 표준편차: {np.std(radii):.3f}")
    
    logger.info("=== 좌표계 중심 제안 ===")
    logger.info("방법 1 - 전체 평균:")
    logger.info(f"  중심점: ({mean_vals[0]:.3f}, {mean_vals[1]:.3f}, {mean_vals[2]:.3f})")
    
    logger.info("방법 2 - XY는 평균, Z는 최소값:")
    logger.info(f"  중심점: ({mean_vals[0]:.3f}, {mean_vals[1]:.3f}, {min_vals[2]:.3f})")
    
    logger.info("방법 3 - XY는 평균, Z는 중간값:")
    logger.info(f"  중심점: ({mean_vals[0]:.3f}, {mean_vals[1]:.3f}, {(min_vals[2] + max_vals[2])/2:.3f})")
    
    # 각 Z 레벨에서의 XY 중심 분석
    z_levels = np.unique(np.round(points[:, 2], 1))
    if len(z_levels) > 10:  # 너무 많으면 샘플링
        z_sample = np.linspace(min_vals[2], max_vals[2], 10)
    else:
        z_sample = z_levels
    
    logger.info("\n=== Z 레벨별 XY 중심 분석 ===")
    xy_centers_by_z = []
    for z_level in z_sample[:5]:  # 상위 5개만 표시
        mask = np.abs(points[:, 2] - z_level) < (max_vals[2] - min_vals[2]) / 20
        if np.sum(mask) > 0:
            level_points = points[mask]
            level_center = np.mean(level_points[:, :2], axis=0)
            xy_centers_by_z.append([z_level, level_center[0], level_center[1]])
            logger.info(f"Z={z_level:.1f}: XY 중심=({level_center[0]:.3f}, {level_center[1]:.3f})")
    
    return {
        'points': points,
        'stats': {
            'min': min_vals,
            'max': max_vals,
            'mean': mean_vals,
            'std': std_vals
        },
        'spring_info': {
            'bottom': spring_bottom,
            'top': spring_top,
            'height': spring_height
        },
        'radius_info': {
            'mean': np.mean(radii),
            'min': np.min(radii),
            'max': np.max(radii),
            'std': np.std(radii)
        }
    }

def visualize_spring(points, analysis_result):
    """스프링 데이터 시각화"""
    if points is None:
        return
    
    fig = plt.figure(figsize=(15, 5))
    
    # 3D 플롯
    ax1 = fig.add_subplot(131, projection='3d')
    ax1.scatter(points[:, 0], points[:, 1], points[:, 2], c=points[:, 2], cmap='viridis', s=1)
    ax1.set_xlabel('X')
    ax1.set_ylabel('Y')
    ax1.set_zlabel('Z')
    ax1.set_title('3D Spring Data')
    
    # XY 평면 플롯
    ax2 = fig.add_subplot(132)
    ax2.scatter(points[:, 0], points[:, 1], c=points[:, 2], cmap='viridis', s=1)
    ax2.set_xlabel('X')
    ax2.set_ylabel('Y')
    ax2.set_title('XY Plane View')
    ax2.axis('equal')
    
    # 중심점 표시
    mean_vals = analysis_result['stats']['mean']
    ax2.plot(mean_vals[0], mean_vals[1], 'r+', markersize=10, markeredgewidth=3, label='Center')
    ax2.legend()
    
    # XZ 평면 플롯
    ax3 = fig.add_subplot(133)
    ax3.scatter(points[:, 0], points[:, 2], c=points[:, 1], cmap='plasma', s=1)
    ax3.set_xlabel('X')
    ax3.set_ylabel('Z')
    ax3.set_title('XZ Plane View')
    
    plt.tight_layout()
    plt.savefig('spring_analysis.png', dpi=150, bbox_inches='tight')
    plt.show()

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filepath = os.path.join(script_dir, 'TK1_FRT_zero-1_251014.igs')
    
    logger.info("IGS 파일 분석 시작...")
    points = parse_igs_points(filepath)
    
    if points is not None:
        analysis_result = analyze_spring_center(points)
        if analysis_result is not None:
            logger.info("\n분석 완료!")
            logger.info("\n=== 결론 ===")
            logger.info("1. IGS 파일로부터 xyz 좌표계 중심을 도출하는 것이 **가능**합니다.")
            logger.info("2. 스프링의 기하학적 중심을 여러 방법으로 계산할 수 있습니다.")
            logger.info("3. rθz 좌표계 설정을 위한 기준점을 제공할 수 있습니다.")
            
            # 시각화
            logger.info("\n시각화 생성 중...")
            visualize_spring(points, analysis_result)
    else:
        logger.error("데이터를 읽을 수 없습니다.")

if __name__ == "__main__":
    main()