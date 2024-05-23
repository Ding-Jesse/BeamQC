import numpy as np
from scipy.optimize import linear_sum_assignment


def calculate_distance_matrix(points1, points2):
    """
    Calculate the Euclidean distance matrix between two lists of points.
    points1 and points2 are lists of tuples in the form (data, (x, y)).
    """
    n, m = len(points1), len(points2)
    distance_matrix = np.zeros((n, m))

    for i in range(n):
        for j in range(m):
            coord1 = np.array(points1[i])
            coord2 = np.array(points2[j])
            distance_matrix[i, j] = np.linalg.norm(coord1 - coord2)

    return distance_matrix


def match_points(points1, points2):
    """
    Match points from two lists using the Hungarian algorithm to minimize total distance.
    points1 and points2 are lists of tuples in the form (data, (x, y)).
    """
    distance_matrix = calculate_distance_matrix(points1, points2)
    row_ind, col_ind = linear_sum_assignment(distance_matrix)
    matched_pairs = [(i, j)
                     for i, j in zip(row_ind, col_ind)]
    total_distance = distance_matrix[row_ind, col_ind].sum()

    return matched_pairs, total_distance


def for_loop_min_match(points1, points2):
    matched_pairs = []
    total_distance = 0

    for i, point1 in enumerate(points1):
        min_distance = float('inf')
        best_match = None
        for j, point2 in enumerate(points2):
            distance = np.linalg.norm(
                np.array(point1) - np.array(point2))
            if distance < min_distance:
                min_distance = distance
                best_match = j
        matched_pairs.append((i, best_match))
        total_distance += min_distance

    return matched_pairs, total_distance
