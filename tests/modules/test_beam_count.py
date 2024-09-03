# test_my_module.py
from src.beam_count import (sort_arrow_line, sort_arrow_to_word, sort_noconcat_line,
                            sort_noconcat_bend, sort_rebar_bend_line, count_tie,
                            add_beam_to_list, combine_beam_boundingbox, break_down_rebar,
                            combine_beam_tie, combine_beam_rebar, compare_line_with_dim,
                            sort_beam)
import pickle
import pytest


@pytest.fixture
def all_test_data():
    pkl_files = [
        r'tests\data\test-data.pkl',
    ]
    test_data_list = []
    for pkl_file in pkl_files:
        with open(pkl_file, 'rb') as f:
            test_data_list.append(pickle.load(f))
    return test_data_list


def test_sort_arrow_line(all_test_data):
    for test_data in all_test_data:
        sort_arrow_line_data = test_data['sort_arrow_line_data']
        inputs = sort_arrow_line_data['inputs']
        expect_outputs = sort_arrow_line_data['outputs']
        coor_to_arrow_dic_output, no_arrow_line_list_output = sort_arrow_line(coor_to_arrow_dic=inputs['coor_to_arrow_dic'],
                                                                              coor_to_bend_rebar_list=inputs[
            'coor_to_bend_rebar_list'],
            coor_to_dim_list=inputs['coor_to_dim_list'],
            coor_to_rebar_list=inputs['coor_to_rebar_list'])

        assert expect_outputs['coor_to_arrow_dic'] == coor_to_arrow_dic_output
        assert expect_outputs['no_arrow_line_list'] == no_arrow_line_list_output


def test_sort_arrow_to_word(all_test_data):
    for test_data in all_test_data:
        case = test_data['sort_arrow_to_word_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']
        coor_to_arrow_dic, head_to_data_dic, tail_to_data_dic = sort_arrow_to_word(
            **inputs)

        assert expect_outputs['coor_to_arrow_dic'] == coor_to_arrow_dic
        assert expect_outputs['head_to_data_dic'] == head_to_data_dic
        assert expect_outputs['tail_to_data_dic'] == tail_to_data_dic


def test_sort_noconcat_line(all_test_data):
    for test_data in all_test_data:
        case = test_data['sort_noconcat_line_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']
        coor_to_rebar_list_straight = sort_noconcat_line(**inputs)

        assert expect_outputs['coor_to_rebar_list_straight'] == coor_to_rebar_list_straight


def test_sort_noconcat_bend(all_test_data):
    for test_data in all_test_data:
        case = test_data['sort_noconcat_bend_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']
        coor_to_bend_rebar_list = sort_noconcat_bend(**inputs)

        assert expect_outputs['coor_to_bend_rebar_list'] == coor_to_bend_rebar_list


def test_sort_rebar_bend_line(all_test_data):
    for test_data in all_test_data:
        case = test_data['sort_rebar_bend_line_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']
        sort_rebar_bend_line(**inputs)

        assert expect_outputs['coor_to_bend_rebar_list'] == inputs['rebar_bend_list']
        assert expect_outputs['coor_to_rebar_list_straight'] == inputs['rebar_line_list']


def test_count_tie(all_test_data):
    for test_data in all_test_data:
        case = test_data['count_tie_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']

        coor_sorted_tie_list = count_tie(**inputs)

        assert expect_outputs['coor_sorted_tie_list'] == coor_sorted_tie_list


def test_add_beam_to_list(all_test_data):
    for test_data in all_test_data:
        case = test_data['add_beam_to_list_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']
        inputs['class_beam_list'] = []
        add_beam_to_list(**inputs)

        for key, item in expect_outputs.items():
            if key in inputs:
                assert expect_outputs[key] == inputs[key]


def test_combine_beam_boundingbox(all_test_data):
    for test_data in all_test_data:
        case = test_data['combine_beam_boundingbox_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']
        combine_beam_boundingbox(**inputs)

        for key, item in expect_outputs.items():
            if key in inputs:
                assert expect_outputs[key] == inputs[key]
        # assert expect_outputs['coor_to_block_list'] == expect_outputs['coor_to_block_list']
        # assert expect_outputs['coor_to_bounding_block_list'] == expect_outputs['coor_to_bounding_block_list']
        # assert expect_outputs['class_beam_list'] == expect_outputs['class_beam_list']
        # assert expect_outputs['coor_to_rc_block_list'] == expect_outputs['coor_to_rc_block_list']


def test_break_down_rebar(all_test_data):
    for test_data in all_test_data:
        case = test_data['break_down_rebar_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']
        break_down_rebar(**inputs)

        for key, item in expect_outputs.items():
            if key in inputs:
                assert expect_outputs[key] == inputs[key]


def test_combine_beam(all_test_data):
    for test_data in all_test_data:
        case = test_data['combine_beam_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']

        combine_beam_tie(**inputs)
        combine_beam_rebar(**inputs)

        for key, item in expect_outputs.items():
            if key in inputs:
                assert expect_outputs[key] == inputs[key]


def test_compare_line_with_dim(all_test_data):
    for test_data in all_test_data:
        case = test_data['compare_line_with_dim_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']

        compare_line_with_dim(**inputs)

        for key, item in expect_outputs.items():
            if key in inputs:
                assert expect_outputs[key] == inputs[key]


def test_sort_beam(all_test_data):
    for test_data in all_test_data:
        case = test_data['sort_beam_data']
        inputs = case['inputs']
        expect_outputs = case['outputs']

        sort_beam(**inputs)

        for key, item in expect_outputs.items():
            if key in inputs:
                assert expect_outputs[key] == inputs[key]


if __name__ == '__main__':
    pkl_files = [
        r'tests\data\test-data.pkl',
    ]
    test_data_list = []
    for pkl_file in pkl_files:
        with open(pkl_file, 'rb') as f:
            data = pickle.load(f)
            test_data_list.append(data)
    test_sort_arrow_to_word(test_data_list)
