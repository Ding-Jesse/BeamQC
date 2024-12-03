import numpy as np
import re
from typing import Literal
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


def convert_mm_to_cm(text):
    # Define a function to convert mm to cm (divide by 10)
    def replace_mm(match):
        # Get the matched number as a string, then convert to float
        mm_value = int(match.group(0))  # Convert string to integer
        # Divide by 10 to convert mm to cm
        cm_value = mm_value / 10.0  # Use float division
        # Return the result formatted to one decimal place if needed
        # Remove trailing zeros and decimal if not necessary
        return f"{cm_value:.1f}".rstrip('0').rstrip('.')

    # Use re.sub to find all numbers and apply the conversion
    result = re.sub(r'\d+', replace_mm, text)
    return result


def find_all_matching_patterns(text, patterns):
    matching_patterns = []
    for pattern in patterns:
        if pattern == '':
            continue
        if re.search(pattern, text):
            matching_patterns.append(pattern)
    return matching_patterns  # Return all matching patterns


def extract_dimensions(text):
    '''
    # Example usage
    text1 = "100x100(cm)"
    text2 = "100.5x100.5(cm)"

    ### Extracting dimensions
    dimensions1 = extract_dimensions(text1) \n
    dimensions2 = extract_dimensions(text2) \n

    print(dimensions1)  # Output: 100x100 \n
    print(dimensions2)  # Output: 100.5x100.5 \n
    '''
    # Updated regular expression to match two sets of digits (with optional decimal) separated by "x"
    match = re.search(r'\d+(\.\d+)?x\d+(\.\d+)?', text)

    # If a match is found, return it; otherwise, return None
    if match:
        return match.group()
    return ''


def inblock(block: tuple, pt: tuple):
    pt_x = pt[0]
    pt_y = pt[1]
    if len(block) == 0:
        return False
    if (pt_x - block[0][0])*(pt_x - block[1][0]) < 0 and (pt_y - block[0][1])*(pt_y - block[1][1]) < 0:
        return True
    return False


def define_beam_type(name_pattern: dict[str, list], beam_name) -> Literal['Grider', 'FB', 'SB', None]:
    if name_pattern:
        for beam_type, patterns in name_pattern.items():
            if beam_type == 'General':
                continue
            match_floor = ''
            match_serial = ''
            match_obj = None

            for pattern in patterns:
                match_obj = re.search(pattern, beam_name)
                if match_obj:
                    match_floor = match_obj.group(1)
                    match_floor = re.sub(r'\(|\)', '', match_floor)  # 去除()
                    match_serial = match_obj.group(2)
                    match_serial.replace(' ', '')  # 去除編號與尺寸的間隔
                    break

            if match_obj:
                return beam_type

    return None

# Custom sorting function


def define_serial_order(item):
    # Regex to match components: prefix, first number, and second number
    match = re.match(r"([A-Za-z]+)(\d+)?(?:-?(\d+))?", item)
    if match:
        prefix = match.group(1)  # The alphabetic prefix (e.g., "B", "G", "GA")
        num1 = int(match.group(2)) if match.group(2) else float(
            'inf')  # First numeric part, or inf if missing
        num2 = int(match.group(3)) if match.group(3) else float(
            'inf')  # Second numeric part, or inf if missing
        # Normalize prefix case for sorting
        return (prefix.lower(), num1, num2)
    # Fallback for unexpected strings
    return (item.lower(), float('inf'), float('inf'))


if __name__ == '__main__':
    pass
