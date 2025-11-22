from __future__ import annotations

import math
from textwrap import dedent

import pytest

from lab_aid.engine import evaluate


def run_e(script: str, inputs: str = "") -> tuple[str | None, str | None, str | None]:
    script_clean = dedent(script).strip()
    inputs_clean = dedent(inputs).strip() if inputs else inputs
    return evaluate("E", script_clean, inputs_clean)


def run_r(script: str, inputs: str) -> tuple[str | None, str | None, str | None]:
    script_clean = dedent(script).strip()
    inputs_clean = dedent(inputs).strip() if inputs else inputs
    return evaluate("R", script_clean, inputs_clean)


def test_e_assign_literal_number():
    assert run_e("this = 42") == ("42", None, None)


def test_e_assign_literal_string():
    assert run_e("this = 'OK'") == ("OK", None, None)


def test_e_assign_variable_number():
    script = """
        value = 7
        this = value
    """
    assert run_e(script) == ("7", None, None)


def test_e_assign_variable_string():
    script = """
        status = 'READY'
        this = status
    """
    assert run_e(script) == ("READY", None, None)


def test_e_hash_lookup_with_unit():
    script = """
        this = #HOLDER.ITEM + #WEIGHT[KG]
    """
    inputs = """
        HOLDER.ITEM=5
        WEIGHT[KG]=3
    """
    assert run_e(script, inputs) == ("8", None, None)


def test_e_basic_arithmetic():
    assert run_e("this = #A * 2", "A=5") == ("10", None, None)


def test_e_parentheses_precedence():
    script = """
        foo = 2
        bar = 3
        this = (foo + #A) * (bar - 1)
    """
    assert run_e(script, "A=4") == ("12", None, None)


def test_e_if_else_branching():
    script = """
        this = 0
        if str_comp(#FLAG, 'YES') eq 0
         this = 1
        else
         this = 2
        end
    """
    assert run_e(script, "FLAG='YES'") == ("1", None, None)


def test_e_comparison_operator():
    script = """
        this = 0
        if #A gt #B
         this = 1
        else
         this = 2
        end
    """
    inputs = """
        A=5
        B=3
    """
    assert run_e(script, inputs) == ("1", None, None)


def test_e_logical_operator():
    script = """
        this = 0
        if str_comp(#A, 'YES') eq 0 and str_comp(#B, 'OK') eq 0
         this = 1
        else
         this = 2
        end
    """
    inputs = """
        A='YES'
        B='OK'
    """
    assert run_e(script, inputs) == ("1", None, None)


def test_e_comparison_and_logical():
    script = """
        this = 0
        if #A gt 5 and (#B le 3 or str_comp(#C, 'OK') eq 0)
         this = 1
        else
         this = 2
        end
    """
    inputs = """
        A=6
        B=4
        C='OK'
    """
    assert run_e(script, inputs) == ("1", None, None)


def test_e_comparison_logical_with_parentheses():
    script = """
        this = 0
        if (#A gt 5 or #B lt 2) and (str_comp(#C, 'NG') ne 0)
         this = 1
        else
         this = 2
        end
    """
    inputs = """
        A=4
        B=1
        C='OK'
    """
    assert run_e(script, inputs) == ("1", None, None)


def test_e_for_loop_accumulates_values():
    script = """
        total = 0
        for I = 1 TO 3
         total = total + I
        next
        this = total
    """
    assert run_e(script) == ("6", None, None)


def test_e_deep_nesting_respects_limits():
    script_ok = """
        total = 0
        for A = 1 TO 2
         for B = 1 TO 2
          total = total + A + B
         next
        next
        this = total
    """
    assert run_e(script_ok) == ("12", None, None)

    script_over = """
        if 1 eq 1
         if 1 eq 1
          if 1 eq 1
           if 1 eq 1
            if 1 eq 1
             if 1 eq 1
              if 1 eq 1
               if 1 eq 1
                if 1 eq 1
                 if 1 eq 1
                  if 1 eq 1
                   this = 1
                  end
                 end
                end
               end
              end
             end
            end
           end
          end
         end
        end
    """
    assert run_e(script_over) == ("エラー", None, None)


def test_e_roundjisb_preserves_trailing_zero():
    assert run_e("this = roundjisb(#A, 2, 1)", "A=1.2") == ("1.20", None, None)


def test_e_floor_positive_value():
    assert run_e("this = floor(#A, 1)", "A=12.39") == ("12.3", None, None)


def test_e_trunc_negative_toward_zero():
    assert run_e("this = trunc(#A, 1)", "A=-1.29") == ("-1.2", None, None)


def test_e_print_outputs():
    script = """
        this = 5
        print(this, 'ED')
    """
    assert run_e(script) == ("5", "ED", None)


def test_e_print2_outputs():
    script = """
        this = 5
        print2(this, 'RP')
    """
    assert run_e(script) == ("5", None, "RP")


def test_e_nested_functions():
    script = "this = roundjisb(floor(#A, 1), 1, 1)"
    assert run_e(script, "A=12.39") == ("12.3", None, None)


def test_e_hash_argument_rejects_varref():
    script = """
        foo = 10
        this = #A + foo
    """
    assert run_e(script, "A=foo") == ("エラー", None, None)


def test_e_sqrt_and_errors():
    assert run_e("this = sqrt(#A)", "A=4") == ("2", None, None)
    assert run_e("this = sqrt(-1)") == ("エラー", None, None)


def test_e_logarithms():
    raw, _, _ = run_e("this = log(#A)", "A=10")
    assert float(raw) == pytest.approx(math.log(10), rel=1e-9)
    assert run_e("this = log(0)") == ("エラー", None, None)
    raw10, _, _ = run_e("this = log10(#A)", "A=100")
    assert float(raw10) == pytest.approx(math.log10(100), rel=1e-9)


def test_e_pow_and_exp():
    assert run_e("this = pow(#A, #B)", "A=2\nB=3") == ("8", None, None)
    raw, _, _ = run_e("this = exp(1)")
    assert float(raw) == pytest.approx(math.e, rel=1e-9)


def test_e_modi_and_modd():
    assert run_e("this = modi(#A)", "A=1234.56") == ("1234", None, None)
    assert run_e("this = modd(#A)", "A=1234.56") == ("0.56", None, None)
    assert run_e("this = modd(#A)", "A=-12.34") == ("0.34", None, None)


def test_e_round_with_unit():
    assert run_e("this = round(12.345, 2, 1, 1)") == ("12.35", None, None)
    assert run_e("this = round(12.345, -1, 1, 1)") == ("10", None, None)


def test_e_round_significant_digits_and_errors():
    assert run_e("this = round(12.345, 3, 0)") == ("12.3", None, None)
    assert run_e("this = round(12.345, 1, 0)") == ("10", None, None)
    assert run_e("this = round(12.345, 1, 1, 5)") == ("12.5", None, None)
    assert run_e("this = round(12.345, 2, 1, 0)") == ("エラー", None, None)
    assert run_e("this = round(12.345, 2, 2)") == ("エラー", None, None)


def test_e_max_min_with_multi_measurements():
    assert run_e("this = max(#A)", "A=1,2,3") == ("3", None, None)
    assert run_e("this = min(#A)", "A=1,2,3") == ("1", None, None)
    assert run_e("this = max(#A, #B)", "A=1,2,3\nB=6,6") == ("6", None, None)
    assert run_e("this = min(#A, #B)", "A=1,2,3\nB=6,6") == ("2", None, None)


def test_e_aggregate_functions():
    assert run_e("this = ave(#A)", "A=1,2,3") == ("2", None, None)
    assert run_e("this = sum(#A)", "A=1,2,3") == ("6", None, None)


def test_e_stdev_and_stdeva():
    raw, _, _ = run_e("this = stdev(#A)", "A=5,5.5,6")
    assert float(raw) == pytest.approx(0.5, rel=1e-9)
    raw_pop, _, _ = run_e("this = stdeva(#A)", "A=5,5.5,6")
    assert float(raw_pop) == pytest.approx(0.4082482904638631, rel=1e-9)
    assert run_e("this = stdev(#A)", "A=5") == ("エラー", None, None)
    assert run_e("this = stdeva(#A)", "A=5") == ("エラー", None, None)


def test_e_string_functions_is_char_strlen():
    assert run_e("this = is_char('ABC')") == ("1", None, None)
    assert run_e("this = is_char(123)") == ("0", None, None)
    assert run_e("this = is_char(#A, 2)", "A='X',123") == ("0", None, None)
    assert run_e("this = strlen(100)") == ("3", None, None)
    assert run_e("this = strlen(#A, 2)", "A='AB','XYZ'") == ("3", None, None)


def test_e_strcat_statement_updates_target():
    script = """
        B = '10'
        strcat(B, '.00')
        this = B
    """
    assert run_e(script) == ("10.00", None, None)

    script = """
        B = #A
        strcat(B, '-end')
        this = B
    """
    assert run_e(script, "A='start'") == ("start-end", None, None)


def test_e_strncpy_statement_uses_shift_jis_bytes():
    script = """
        strncpy(B, 'ABCDEFG', 2, 3)
        this = B
    """
    assert run_e(script) == ("BCD", None, None)

    script = """
        strncpy(B, 'ABCDEFG', 5)
        this = B
    """
    assert run_e(script) == ("EFG", None, None)

    script = """
        strncpy(B, ' あいうえお ', 3, 4)
        this = B
    """
    assert run_e(script) == ("いう", None, None)


def test_e_string_functions_empty_and_space():
    assert run_e("this = isempty('')") == ("1", None, None)
    assert run_e("this = isempty('text')") == ("0", None, None)
    assert run_e("this = isempty(#A, 2)", "A='', 'text'") == ("0", None, None)
    assert run_e("this = isspace(' ')") == ("1", None, None)
    assert run_e("this = isspace('abc')") == ("0", None, None)
    assert run_e("this = isspace(#A, 2)", "A=' ', ' \t'") == ("1", None, None)


def test_e_multi_unit_and_holder_inputs():
    script = """
        this = #HOLDER.ITEM + #WEIGHT[KG] + #TEMP[C]
    """
    inputs = """
        HOLDER.ITEM=5
        WEIGHT[KG]=3
        TEMP[C]=2
    """
    assert run_e(script, inputs) == ("10", None, None)


def test_r_uses_input_value_on_rhs():
    assert run_r("this = this + 3", "7") == (None, "10", "10")


def test_r_varref_defaults_to_zero_before_execution():
    assert run_r("this = this + 1", "") == (None, "エラー", "エラー")


def test_r_rejects_hash_variables():
    assert run_r("this = #A", "1") == (None, "エラー", "エラー")


def test_r_print_overrides_this_value():
    script = """
        this = this + 2
        print(this, 'ED')
    """
    assert run_r(script, "3") == (None, "ED", "ED")


def test_r_print2_overrides_reported_only():
    script = """
        print(this, 'ED')
        print2(this, 'RP')
    """
    assert run_r(script, "4") == (None, "ED", "RP")


def test_r_roundjisb_uses_input_as_this():
    assert run_r("this = roundjisb(this, 2, 1)", "1.205") == (None, "1.21", "1.21")


def test_r_literal_echo_when_no_this_or_print():
    script = """
        this = this + 0
    """
    assert run_r(script, "0") == (None, "0", "0")


def test_r_allows_hash_in_literals_and_comments():
    script = """
        rem コメント内の #HASH は許容される
        note = 'value #1'
        this = this + 1
    """
    assert run_r(script, "2") == (None, "3", "3")


def test_r_print2_condition_false_keeps_literal():
    script = """
        if this lt 0.23
         print2(this, '不検出')
        end
    """
    assert run_r(script, "0.36") == (None, "0.36", "0.36")


def test_r_rhs_this_always_uses_input_value():
    script = """
        this = this + 2
        this = this * 2
    """
    # 2 行目でも RHS の this は入力値 (3) を参照するため、(3 * 2) となる
    assert run_r(script, "3") == (None, "6", "6")


def test_r_nested_for_and_if_structure():
    script = """
        total = 0
        for I = 1 TO 2
         if I eq 1
          total = total + this
         else
          total = total + 2
         end
        next
        this = total
    """
    # 1 回目は入力値 (3) を加算、2 回目は 2 を加算して合計 5
    assert run_r(script, "3") == (None, "5", "5")


def test_r_comments_and_literals_ignore_hash_and_this():
    script = """
        rem コメント内の #TEMP や this は無視される
        note = 'value #1 with this keyword'
        this = this + 2
    """
    assert run_r(script, "3") == (None, "5", "5")


def test_e_nested_for_if_structure():
    script = """
        sum = 0
        for I = 1 TO 2
         if I eq 1
          sum = sum + 2
         else
          sum = sum + 3
         end
        next
        this = sum
    """
    assert run_e(script) == ("5", None, None)


def test_e_roundjisb_invalid_argument_count():
    assert run_e("this = roundjisb(#A, 1)", "A=10") == ("エラー", None, None)


def test_e_roundjisb_preserves_fixed_decimals_through_variables():
    script = """
        a = roundjisb(12.300, 2, 1)
        this = a
    """
    assert run_e(script) == ("12.30", None, None)


def test_e_roundjisb_preserves_fixed_decimals_direct_assignment():
    assert run_e("this = roundjisb(12.300, 3, 1)") == ("12.300", None, None)


def test_e_roundjisb_nested_with_floor_keeps_trailing_zero():
    script = """
        floored = floor(12.3456, 2)
        this = roundjisb(floored, 3, 1)
    """
    assert run_e(script) == ("12.340", None, None)


def test_e_roundjisb_floor_inline_expression_preserves_decimals():
    script = "this = roundjisb(floor(0.109, 2), 2, 1)"
    assert run_e(script) == ("0.10", None, None)


# def test_e_for_loop_iteration_limits():
#     # 反復が制限内なら実行できる
#     script_limit = """
#         total = 0
#         for I = 1 TO 1000000
#          if I le 3
#           total = total + 1
#          end
#         next
#         this = total
#     """
#     assert run_e(script_limit) == ("3", None, None)

#     # 上限を超える for ループはエラー
#     script_over = """
#         total = 0
#         for I = 1 TO 1000001
#          total = total + 1
#         next
#         this = total
#     """
#     assert run_e(script_over) == ("エラー", None, None)


def test_e_floor_invalid_argument_type():
    assert run_e("this = floor('text', 1)", "") == ("エラー", None, None)


def test_e_trunc_invalid_argument_type():
    assert run_e("this = trunc(#A, 'bad')", "A=5") == ("エラー", None, None)


def test_e_hash_item_missing_raises_error():
    assert run_e("this = #MISSING", "") == ("エラー", None, None)


def test_e_strncpy_invalid_start_and_length():
    assert run_e("this = strncpy('', 'ABC', 0)") == ("エラー", None, None)
    assert run_e("this = strncpy('', 'ABC', 2, -1)") == ("エラー", None, None)


def test_e_str_comp_uses_multi_measurement_indices():
    script = "this = str_comp(#A, #A, 1, 2)"
    assert run_e(script, "A='OK','NG'") == ("1", None, None)


def test_e_str_comp_default_first_measurement():
    script = "this = str_comp(#A, 'OK')"
    assert run_e(script, "A='OK','NG'") == ("0", None, None)


def test_e_str_comp_rejects_literal_left_operand():
    assert run_e("this = str_comp('A', 'B')") == ("エラー", None, None)


def test_e_str_comp_requires_variable_index_in_three_arg_form():
    script = "this = str_comp(#A, 1, 'OK')"
    assert run_e(script, "A='OK','NG'") == ("エラー", None, None)

    script = """
        idx = 2
        this = str_comp(#A, idx, 'NG')
    """
    assert run_e(script, "A='OK','NG'") == ("0", None, None)


def test_e_str_comp_requires_string_literal_in_three_arg_form():
    script = """
        idx = 1
        this = str_comp(#A, idx, #B)
    """
    assert run_e(script, "A='OK'\nB='NG'") == ("エラー", None, None)


def test_e_requires_this_assignment():
    assert run_e("foo = 1") == ("エラー", None, None)


def test_e_missing_end_reports_error():
    script = """
        if 1 eq 1
         this = 1
    """
    assert run_e(script) == ("エラー", None, None)


def test_e_for_with_zero_step_reports_error():
    script = """
        this = 0
        for I = 1 TO 3 STEP 0
         this = this + 1
        next
    """
    assert run_e(script) == ("エラー", None, None)


def test_e_for_with_negative_step():
    script = """
        total = 0
        for I = 3 TO 1 STEP -1
         total = total + I
        next
        this = total
    """
    assert run_e(script) == ("6", None, None)


def test_r_print2_preserves_roundjisb_formatting():
    script = """
        this=roundjisb(this,3,1)
        a=roundjisb(this,2,1)
        strcat(a,'')
        print2(this,a)
    """
    assert run_r(script, "0") == (None, "0.000", "0.00")
